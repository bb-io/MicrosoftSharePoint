using Apps.MicrosoftSharePoint.Dtos;
using Blackbird.Applications.Sdk.Common.Authentication;
using Apps.MicrosoftSharePoint.Extensions;
using RestSharp;

namespace Apps.MicrosoftSharePoint;

public class MicrosoftSharePointClient : RestClient
{
    public MicrosoftSharePointClient(IEnumerable<AuthenticationCredentialsProvider> authenticationCredentialsProviders) 
        : base(new RestClientOptions
        {
            ThrowOnAnyError = false, BaseUrl = GetBaseUrl(authenticationCredentialsProviders)
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
        var response = await ExecuteAsync(request);
        
        if (response.IsSuccessful)
            return response;

        throw ConfigureErrorException(response.Content);
    }

    private Exception ConfigureErrorException(string responseContent)
    {
        var error = responseContent.DeserializeObject<ErrorDto>();
        return new($"{error.Error.Code}: {error.Error.Message}");
    }
}