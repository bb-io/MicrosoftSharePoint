using Blackbird.Applications.Sdk.Common.Invocation;

namespace Apps.MicrosoftSharePoint.Webhooks.Handlers;

public class DriveWebhookHandler : BaseWebhookHandler
{
    const string SubscriptionEvent = "updated"; // the only event type supported for drive items
    
    public DriveWebhookHandler(InvocationContext invocationContext) 
        : base(invocationContext, SubscriptionEvent) { }

    protected override string GetResource()
    {
        var siteId = InvocationContext.AuthenticationCredentialsProviders.First(p => p.KeyName == "SiteId").Value;
        return $"/sites/{siteId}/drive/root";
    }
}