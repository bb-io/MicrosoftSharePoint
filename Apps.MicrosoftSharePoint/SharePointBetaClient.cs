using Apps.MicrosoftSharePoint.Dtos;
using Blackbird.Applications.Sdk.Common.Authentication;
using Apps.MicrosoftSharePoint.Extensions;
using RestSharp;
using Blackbird.Applications.Sdk.Common.Exceptions;
using System.Globalization;
using System.Net;

namespace Apps.MicrosoftSharePoint;

public class SharePointBetaClient : RestClient
{
    private const int MaxRetries = 5;
    private const int InitialDelayMs = 1000;
    public SharePointBetaClient(IEnumerable<AuthenticationCredentialsProvider> authenticationCredentialsProviders) 
        : base(new RestClientOptions
        {
            ThrowOnAnyError = false, BaseUrl = GetBaseUrl(authenticationCredentialsProviders),
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
        int delay = InitialDelayMs;
        RestResponse? response = null;

        for (int attempt = 1; attempt <= MaxRetries; attempt++)
        {
            response = await ExecuteAsync(request);

            if (response.IsSuccessful)
                return response;

            if (attempt < MaxRetries &&
                (response.StatusCode == HttpStatusCode.InternalServerError ||
                 response.StatusCode == HttpStatusCode.ServiceUnavailable || 
                 response.StatusCode == HttpStatusCode.BadRequest ||
                 response.StatusCode == HttpStatusCode.TooManyRequests))
            {
                await Task.Delay(delay);
                delay *= 2;
                continue;
            }
            break;
        }
        throw ConfigureErrorException(response?.Content ?? string.Empty);
    }

    private Exception ConfigureErrorException(string responseContent)
    {
        var error = responseContent.DeserializeObject<ErrorDto>();

        if (error.Error.Message.Contains("InternalServerError", StringComparison.OrdinalIgnoreCase) == true)
        {
            return new PluginApplicationException("An internal server error occurred. Please implement a retry policy and try again.");
        }
        if (error.Error.Message.Contains("ServiceUnavailable", StringComparison.OrdinalIgnoreCase) == true)
        {
            return new PluginApplicationException("Currently the Sharepoint service is not available. Please check your credentials or implement a retry policy and try again.");
        }
        if (error.Error.Code?.Equals("TooManyRequests", StringComparison.OrdinalIgnoreCase) == true)
        {
            return new PluginApplicationException("Too many requests. Please wait and try again later.");
        }
        if (error.Error.Message.Contains("The resource could not be found", StringComparison.OrdinalIgnoreCase) == true)
        {
            return new PluginMisconfigurationException(
                "The resource URL could not be found. Try to adjust your connection URL: " +
                "https://docs.blackbird.io/apps/microsoft-sharepoint/#how-to-find-site-url"
            );
        }

        return new PluginApplicationException(error.Error.Message);
    }
}