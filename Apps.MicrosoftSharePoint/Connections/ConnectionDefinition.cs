using Apps.MicrosoftSharePoint.Dtos;
using Blackbird.Applications.Sdk.Common.Authentication;
using Blackbird.Applications.Sdk.Common.Connections;
using Apps.MicrosoftSharePoint.Extensions;
using RestSharp;

namespace Apps.MicrosoftSharePoint.Connections;

public class ConnectionDefinition : IConnectionDefinition
{
    public IEnumerable<ConnectionPropertyGroup> ConnectionPropertyGroups => new List<ConnectionPropertyGroup>
    {
        new()
        {
            Name = "OAuth",
            AuthenticationType = ConnectionAuthenticationType.OAuth2,
            ConnectionProperties = new List<ConnectionProperty>
            {
                new("Site name")
            }
        },
    };

    public IEnumerable<AuthenticationCredentialsProvider> CreateAuthorizationCredentialsProviders(
        Dictionary<string, string> values)
    {
        var token = values.First(v => v.Key == "access_token").Value;
        yield return new AuthenticationCredentialsProvider(
            "Authorization",
            $"Bearer {token}"
        );
        
        var siteDisplayName = values.First(v => v.Key == "Site name").Value.Trim();
        var siteId = GetSiteId(token, siteDisplayName);
        
        yield return new AuthenticationCredentialsProvider(
            "SiteId",
            siteId
        );
    }

    private string? GetSiteId(string accessToken, string siteDisplayName)
    {
        var client = new SharePointClient();
        var endpoint = "/sites?search=*";
        string siteId;

        do
        {
            var request = new RestRequest(endpoint);
            request.AddHeader("Authorization", $"Bearer {accessToken}");
            var response = client.Get(request);
            var resultSites = response.Content.DeserializeObject<ListWrapper<SiteDto>>();
            siteId = resultSites.Value.FirstOrDefault(site => site.DisplayName == siteDisplayName)?.Id;

            if (siteId != null)
                break;

            endpoint = resultSites.ODataNextLink?.Split("v1.0")[^1];
        } while (endpoint != null);

        return siteId;
    }
}