using System.Net;
using Apps.MicrosoftSharePoint.Dtos;
using Apps.MicrosoftSharePoint.Extensions;
using Blackbird.Applications.Sdk.Common.Authentication;
using Blackbird.Applications.Sdk.Common.Exceptions;
using Polly.Retry;
using RestSharp;

namespace Apps.MicrosoftSharePoint.Api;

public class SharePointBetaClient : RestClient
{
    private const int RetryCount = 8;

    private readonly AsyncRetryPolicy<RestResponse> _retryPolicy = SharePointRetryPolicy.Create(RetryCount);

    public SharePointBetaClient(IEnumerable<AuthenticationCredentialsProvider> authenticationCredentialsProviders) 
        : base(new RestClientOptions
        {
            ThrowOnAnyError = false, 
            BaseUrl = GetBaseUrl(authenticationCredentialsProviders),
            Timeout = TimeSpan.FromMinutes(5),
        }) { }

    private static Uri GetBaseUrl(IEnumerable<AuthenticationCredentialsProvider> authenticationCredentialsProviders)
    {
        var siteId = authenticationCredentialsProviders.First(p => p.KeyName == "SiteId").Value;
        return new($"https://graph.microsoft.com/beta/sites/{siteId}");
    }

    public async Task<T> ExecuteWithHandling<T>(RestRequest request)
    {
        var response = await ExecuteWithHandling(request);
        return response.Content.DeserializeObject<T>();
    }
    
    public async Task<RestResponse> ExecuteWithHandling(RestRequest request)
    {
        var response = await _retryPolicy.ExecuteAsync(() => ExecuteAsync(request));

        if (response.IsSuccessful)
            return response;

        throw ConfigureErrorException(response);
    }

    private Exception ConfigureErrorException(RestResponse? response)
    {
        if (response == null)
        {
            return new PluginApplicationException("Request failed: No response received from SharePoint.");
        }
        
        if (response.StatusCode == HttpStatusCode.ServiceUnavailable)
        {
            return new PluginApplicationException("SharePoint service is currently unavailable. All retry attempts failed. Please try again later.");
        }

        if (response.StatusCode == HttpStatusCode.InternalServerError)
        {
            return new PluginApplicationException("SharePoint internal server error occurred. All retry attempts failed. Please try again later.");
        }

        if (response.StatusCode == HttpStatusCode.TooManyRequests)
        {
            return new PluginApplicationException("Too many requests to SharePoint. All retry attempts failed. Please wait and try again later.");
        }
        
        try
        {
            var responseContent = response.Content ?? string.Empty;
            if (string.IsNullOrWhiteSpace(responseContent))
            {
                return new PluginApplicationException($"Request failed with status code {response.StatusCode}. No error details provided.");
            }

            var error = responseContent.DeserializeObject<ErrorDto>();

            if (error?.Error == null)
            {
                return new PluginApplicationException($"Request failed with status code {response.StatusCode}. Response: {responseContent}");
            }

            var errorMessage = error.Error.Message ?? "Unknown error";
            var errorCode = error.Error.Code ?? response.StatusCode.ToString();

            if (errorMessage.Contains("InternalServerError", StringComparison.OrdinalIgnoreCase))
            {
                return new PluginApplicationException("SharePoint internal server error occurred. All retry attempts failed. Please try again later.");
            }

            if (errorMessage.Contains("ServiceUnavailable", StringComparison.OrdinalIgnoreCase))
            {
                return new PluginApplicationException("SharePoint service is currently unavailable. All retry attempts failed. Please try again later.");
            }

            if (errorCode.Equals("TooManyRequests", StringComparison.OrdinalIgnoreCase))
            {
                return new PluginApplicationException("Too many requests to SharePoint. All retry attempts failed. Please wait and try again later.");
            }

            if (errorMessage.Contains("The resource could not be found", StringComparison.OrdinalIgnoreCase))
            {
                return new PluginMisconfigurationException(
                    "The resource URL could not be found. Try to adjust your connection URL: " +
                    "https://docs.blackbird.io/apps/microsoft-sharepoint/#how-to-find-site-url"
                );
            }

            return new PluginApplicationException($"SharePoint error: {errorMessage}");
        }
        catch (Exception ex)
        {
            return new PluginApplicationException($"Request failed with status code {response.StatusCode}. Error parsing response: {ex.Message}");
        }
    }
}