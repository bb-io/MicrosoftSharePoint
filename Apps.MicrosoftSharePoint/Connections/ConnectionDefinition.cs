using Apps.MicrosoftSharePoint.Dtos;
using Blackbird.Applications.Sdk.Common.Authentication;
using Blackbird.Applications.Sdk.Common.Connections;
using Apps.MicrosoftSharePoint.Extensions;
using RestSharp;
using Blackbird.Applications.Sdk.Common.Exceptions;
using System.Net;

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
        
        var siteDisplayName = values.First(v => v.Key == "Site URL").Value.Trim();
        var siteId = GetSiteId(token, siteDisplayName);
        
        yield return new AuthenticationCredentialsProvider(
            "SiteId",
            siteId
        );
    }

    public string? GetSiteId(string accessToken, string siteUrl)
    {
        if (!Uri.TryCreate(siteUrl, UriKind.Absolute, out var uri))
            throw new PluginMisconfigurationException($"Invalid SharePoint site URL: {siteUrl}");

        var host = uri.Host;

        var serverRelativePath = uri.AbsolutePath;
        serverRelativePath = "/" + serverRelativePath.Trim('/');

        var endpoint = serverRelativePath == "/"
           ? $"/sites/{host}"
           : $"/sites/{host}:{serverRelativePath}";

        var client = new SharePointClient();
        var request = new RestRequest(endpoint);
        request.AddHeader("Authorization", $"Bearer {accessToken}");

        var response = client.Get(request);
        if (response.StatusCode != HttpStatusCode.OK || string.IsNullOrWhiteSpace(response.Content))
            throw new PluginApplicationException(
                $"Failed to resolve site by URL '{siteUrl}'. " +
                $"Status: {(int)response.StatusCode} {response.StatusDescription}. Body: {response.Content}");

        var site = response.Content.DeserializeObject<SiteDto>()
                   ?? throw new PluginApplicationException($"Unexpected response while resolving site '{siteUrl}'.");

        if (string.IsNullOrWhiteSpace(site.Id))
            throw new PluginApplicationException($"Resolved site has empty Id. URL: {siteUrl}");

        return site.Id!;
    }
}