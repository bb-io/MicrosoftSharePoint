using Apps.MicrosoftSharePoint.Helper;
using Apps.MicrosoftSharePoint.Models.Identifiers;
using Blackbird.Applications.Sdk.Common.Webhooks;
using Blackbird.Applications.Sdk.Common.Invocation;

namespace Apps.MicrosoftSharePoint.Webhooks.Handlers;

public class DriveWebhookHandler(InvocationContext invocationContext, [WebhookParameter] FolderIdentifier folder) 
    : BaseWebhookHandler(invocationContext, SubscriptionEvent)
{
    private const string SubscriptionEvent = "updated";
    private readonly FolderIdentifier _folder = folder;

    protected override string GetResource()
    {
        var siteId = InvocationContext.AuthenticationCredentialsProviders.First(p => p.KeyName == "SiteId").Value;
        var location = ItemIdParser.Parse(_folder?.FolderId);
        if (location.IsDefaultDrive)
            return $"/sites/{siteId}/drive/root";
        else
            return $"/sites/{siteId}/drives/{location.DriveId}/root";
    }
}