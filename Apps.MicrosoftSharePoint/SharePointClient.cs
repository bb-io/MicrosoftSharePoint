using Apps.MicrosoftSharePoint.Dtos;
using Apps.MicrosoftSharePoint.Extensions;
using RestSharp;
using Blackbird.Applications.Sdk.Common.Exceptions;
using System.Net;

namespace Apps.MicrosoftSharePoint;

public class SharePointClient : RestClient
{
    private const int MaxRetries = 5;
    private const int InitialDelayMs = 1000;

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
                 response.StatusCode == HttpStatusCode.BadRequest))
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

        if (error.Error.Code?.Equals("InternalServerError", StringComparison.OrdinalIgnoreCase) == true)
        {
            return new PluginApplicationException("An internal server error occurred. Please implement a retry policy and try again.");
        }
        return new PluginApplicationException($"{error.Error.Code} - {error.Error.Message}");
    }
}