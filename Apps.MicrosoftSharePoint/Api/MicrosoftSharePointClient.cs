using Apps.MicrosoftSharePoint.Models.Dtos;
using Blackbird.Applications.Sdk.Common.Authentication;
using Blackbird.Applications.Sdk.Utils.RestSharp;
using Newtonsoft.Json;
using RestSharp;

namespace Apps.MicrosoftSharePoint.Api;

public class MicrosoftSharePointClient : BlackBirdRestClient
{
    protected override JsonSerializerSettings JsonSettings =>
        new()
        {
            MissingMemberHandling = MissingMemberHandling.Ignore, DateTimeZoneHandling = DateTimeZoneHandling.Local
        };

    public MicrosoftSharePointClient(IEnumerable<AuthenticationCredentialsProvider> authenticationCredentialsProviders,
        bool isBeta = false)
        : base(new RestClientOptions
        {
            ThrowOnAnyError = false, BaseUrl = GetBaseUrl(authenticationCredentialsProviders, isBeta)
        })
    {
        this.AddDefaultHeader("Authorization",
            authenticationCredentialsProviders.First(p => p.KeyName == "Authorization").Value);
    }

    private static Uri GetBaseUrl(IEnumerable<AuthenticationCredentialsProvider> authenticationCredentialsProviders, 
        bool isBeta)
    {
        var siteId = authenticationCredentialsProviders.First(p => p.KeyName == "SiteId").Value;
        
        if (isBeta)
            return new($"https://graph.microsoft.com/beta/sites/{siteId}");
        
        return new($"https://graph.microsoft.com/v1.0/sites/{siteId}");
    }
    
    protected override Exception ConfigureErrorException(RestResponse response)
    {
        var error = JsonConvert.DeserializeObject<ErrorDto>(response.Content, JsonSettings);
        return new($"{error.Error.Code}: {error.Error.Message}");
    }
}