using Apps.MicrosoftSharePoint.Dtos;
using Apps.MicrosoftSharePoint.Extensions;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Authentication;
using Blackbird.Applications.Sdk.Common.Invocation;
using Blackbird.Applications.Sdk.Common.Webhooks;
using Microsoft.AspNetCore.WebUtilities;
using RestSharp;

namespace Apps.MicrosoftSharePoint.Webhooks.Handlers;

public abstract class BaseWebhookHandler : BaseInvocable, IWebhookEventHandler, IAsyncRenewableWebhookEventHandler
{
    private string BridgeWebhooksUrl = "";
    
    private readonly string _subscriptionEvent;
    
    private string Resource => GetResource();

    protected BaseWebhookHandler(InvocationContext invocationContext, string subscriptionEvent)
        : base(invocationContext)
    {
        _subscriptionEvent = subscriptionEvent;
        BridgeWebhooksUrl = InvocationContext.UriInfo.BridgeServiceUrl.ToString().TrimEnd('/') + $"/webhooks/{ApplicationConstants.AppName}";
    }

    protected abstract string GetResource();
    
    public async Task SubscribeAsync(IEnumerable<AuthenticationCredentialsProvider> authenticationCredentialsProviders,
        Dictionary<string, string> values)
    {
        var sharePointClient = new RestClient(new RestClientOptions("https://graph.microsoft.com/v1.0"));
        var targetSubscription = await GetTargetSubscription(authenticationCredentialsProviders, sharePointClient);
        var bridgeService = new BridgeService(InvocationContext.UriInfo.BridgeServiceUrl.ToString().TrimEnd('/'));
        string subscriptionId;
        
        if (targetSubscription is null)
        {
            var createSubscriptionRequest = new MicrosoftSharePointRequest("/subscriptions", Method.Post,
                authenticationCredentialsProviders);
            createSubscriptionRequest.AddJsonBody(new
            {
                ChangeType = _subscriptionEvent,
                NotificationUrl = BridgeWebhooksUrl,
                Resource = Resource,
                ExpirationDateTime = (DateTime.Now + TimeSpan.FromMinutes(40000)).ToString("O"),
                ClientState = ApplicationConstants.SharePointClientState
            });
            
            var response = await sharePointClient.ExecuteAsync(createSubscriptionRequest);
            var subscription = response.Content.DeserializeObject<SubscriptionDto>();
            subscriptionId = subscription.Id;
            
            var deltaRequest = new MicrosoftSharePointRequest($"{Resource}/delta", Method.Get, 
                authenticationCredentialsProviders); 
            response = await sharePointClient.ExecuteAsync(deltaRequest);
            var result = response.Content.DeserializeObject<ListWrapper<object>>();
            
            while (result.ODataNextLink != null)
            {
                var endpoint = result.ODataNextLink?.Split("v1.0")[1];
                deltaRequest = new MicrosoftSharePointRequest(endpoint, Method.Get, authenticationCredentialsProviders);
                response = await sharePointClient.ExecuteAsync(deltaRequest);
                result = response.Content.DeserializeObject<ListWrapper<object>>();
            }
            
            var deltaToken = QueryHelpers.ParseQuery(result.ODataDeltaLink!.Split("?")[1])["token"];
            await bridgeService.StoreValue(subscriptionId, deltaToken);
        }
        else
            subscriptionId = targetSubscription.Id;
        
        await bridgeService.Subscribe(values["payloadUrl"], subscriptionId, _subscriptionEvent);
    }

    public async Task UnsubscribeAsync(IEnumerable<AuthenticationCredentialsProvider> authenticationCredentialsProviders,
        Dictionary<string, string> values)
    {
        var sharePointClient = new RestClient(new RestClientOptions("https://graph.microsoft.com/v1.0"));
        var targetSubscription = await GetTargetSubscription(authenticationCredentialsProviders, sharePointClient);
        var subscriptionId = targetSubscription.Id;
        
        var bridgeService = new BridgeService(InvocationContext.UriInfo.BridgeServiceUrl.ToString().TrimEnd('/'));
        var webhooksLeft = await bridgeService.Unsubscribe(values["payloadUrl"], subscriptionId, _subscriptionEvent);

        if (webhooksLeft == 0)
        {
            await bridgeService.DeleteValue(subscriptionId);
            var deleteSubscriptionRequest = new MicrosoftSharePointRequest($"/subscriptions/{subscriptionId}", 
                Method.Delete, authenticationCredentialsProviders);
            await sharePointClient.ExecuteAsync(deleteSubscriptionRequest);
        }
    }

    [Period(39995)]
    public async Task RenewSubscription(IEnumerable<AuthenticationCredentialsProvider> authenticationCredentialsProviders,
        Dictionary<string, string> values)
    {
        var sharePointClient = new RestClient(new RestClientOptions("https://graph.microsoft.com/v1.0"));
        var targetSubscription = await GetTargetSubscription(authenticationCredentialsProviders, sharePointClient);
        var updateSubscriptionRequest = new MicrosoftSharePointRequest($"/subscriptions/{targetSubscription.Id}", 
            Method.Patch, authenticationCredentialsProviders);
        updateSubscriptionRequest.AddJsonBody(new
        {
            ExpirationDateTime = (DateTime.Now + TimeSpan.FromMinutes(40000)).ToString("O")
        });
        await sharePointClient.ExecuteAsync(updateSubscriptionRequest);
    }

    private async Task<SubscriptionDto?> GetTargetSubscription(
        IEnumerable<AuthenticationCredentialsProvider> authenticationCredentialsProviders, 
        RestClient sharePointClient)
    {
        var subscriptionsRequest = new MicrosoftSharePointRequest("/subscriptions", Method.Get, 
            authenticationCredentialsProviders);
        var response = await sharePointClient.ExecuteAsync(subscriptionsRequest);
        var subscriptions = response.Content.DeserializeObject<SubscriptionWrapper>().Value;
        var targetSubscription = subscriptions.FirstOrDefault(s => s.Resource == Resource 
                                                                   && s.ChangeType == _subscriptionEvent
                                                                   && s.NotificationUrl == BridgeWebhooksUrl);
        return targetSubscription;
    }
}