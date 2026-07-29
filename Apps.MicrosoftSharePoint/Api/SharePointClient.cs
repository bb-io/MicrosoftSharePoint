using System.Net;
using Apps.MicrosoftSharePoint.Dtos;
using Apps.MicrosoftSharePoint.Extensions;
using Blackbird.Applications.Sdk.Common.Exceptions;
using Polly.Retry;
using RestSharp;

namespace Apps.MicrosoftSharePoint.Api;

public class SharePointClient : RestClient
{
    private const int RetryCount = 8;

    private readonly AsyncRetryPolicy<RestResponse> _retryPolicy = SharePointRetryPolicy.Create(RetryCount);

    public SharePointClient(string baseUrl)
        : base(new RestClientOptions
        {
            ThrowOnAnyError = false, BaseUrl = new Uri(baseUrl)
        }) { }

    public SharePointClient()
        : base(new RestClientOptions
        {
            ThrowOnAnyError = false,
            BaseUrl = new Uri("https://graph.microsoft.com/v1.0")
        })
    { }

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
            return new PluginApplicationException("Request failed: No response received from SharePoint.");

        if (response.StatusCode == HttpStatusCode.TooManyRequests)
            return new PluginApplicationException(
                "Too many requests to SharePoint. All retry attempts failed. Please wait and try again later.");

        var responseContent = response.Content ?? string.Empty;

        if (string.IsNullOrWhiteSpace(responseContent))
            return new PluginApplicationException(
                $"Request failed with status code {response.StatusCode}. No error details provided.");

        ErrorDto? error;
        try
        {
            error = responseContent.DeserializeObject<ErrorDto>();
        }
        catch (Exception)
        {
            error = null;
        }

        if (error?.Error == null)
            return new PluginApplicationException(
                $"Request failed with status code {response.StatusCode}. Response: {responseContent}");

        var errorMessage = error.Error.Message;

        if ((errorMessage?.Contains("Internal Server Error", StringComparison.OrdinalIgnoreCase) ?? false) || (errorMessage?.Contains("InternalServerError", StringComparison.OrdinalIgnoreCase) ?? false))
        {
            return new PluginApplicationException("An internal server error occurred. All retry attempts failed. Please try again later.");
        }

        if ((errorMessage?.Contains("Service Unavailable", StringComparison.OrdinalIgnoreCase) ?? false) || (errorMessage?.Contains("ServiceUnavailable", StringComparison.OrdinalIgnoreCase) ?? false))
        {
            return new PluginApplicationException("Server service unavailable error occurred. All retry attempts failed. Please try again later.");
        }
        return new PluginApplicationException($"{error.Error.Code} - {errorMessage}");
    }
}
