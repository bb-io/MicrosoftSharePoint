using Apps.MicrosoftSharePoint.Models.Entities;
using Blackbird.Applications.Sdk.Common.Authentication;
using RestSharp;

namespace Apps.MicrosoftSharePoint.Helper;

public static class DriveHelper
{
    public static async Task<DriveEntity> GetDefaultDrive(
        IEnumerable<AuthenticationCredentialsProvider> creds,
        SharePointBetaClient client)
    {
        var siteId = creds.First(x => x.KeyName == "SiteId").Value;

        var request = new SharePointRequest($"/sites/{siteId}/drive", Method.Get, creds);
        return await client.ExecuteWithHandling<DriveEntity>(request);
    }
}
