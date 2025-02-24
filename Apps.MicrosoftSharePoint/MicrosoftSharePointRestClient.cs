using Apps.MicrosoftSharePoint.Dtos;
using Apps.MicrosoftSharePoint.Extensions;
using RestSharp;
using Blackbird.Applications.Sdk.Common.Exceptions;

namespace Apps.MicrosoftSharePoint;

public class MicrosoftSharePointRestClient : RestClient
{
    public MicrosoftSharePointRestClient(string baseUrl) 
        : base(new RestClientOptions
        {
            ThrowOnAnyError = false, BaseUrl = new Uri(baseUrl)
        }) { }

    public MicrosoftSharePointRestClient()
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
        var response = await ExecuteAsync(request);
        
        if (response.IsSuccessful)
            return response;

        throw ConfigureErrorException(response.Content);
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