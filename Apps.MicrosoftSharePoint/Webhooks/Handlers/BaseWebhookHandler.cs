using Apps.MicrosoftSharePoint.Api;
using Apps.MicrosoftSharePoint.Extensions;
using Apps.MicrosoftSharePoint.Models.Dtos;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Authentication;
using Blackbird.Applications.Sdk.Common.Invocation;
using Blackbird.Applications.Sdk.Common.Webhooks;
using Microsoft.AspNetCore.WebUtilities;
using RestSharp;

namespace Apps.MicrosoftSharePoint.Webhooks.Handlers;

public abstract class BaseWebhookHandler : BaseInvocable, IWebhookEventHandler, IAsyncRenewableWebhookEventHandler
{
    private readonly string _bridgeWebhooksUrl;
    private readonly string _bridgeUrl;
    private readonly string _subscriptionEvent;
    private readonly RestClient _graphClient;
    
    private string Resource => GetResource();
    
    protected BaseWebhookHandler(InvocationContext invocationContext, string subscriptionEvent)
        : base(invocationContext)
    {
        _subscriptionEvent = subscriptionEvent;
        _bridgeUrl = InvocationContext.UriInfo.BridgeServiceUrl.ToString().TrimEnd('/');
        _bridgeWebhooksUrl = _bridgeUrl + $"/webhooks/{ApplicationConstants.AppName}";
        _graphClient = new RestClient(new RestClientOptions("https://graph.microsoft.com/v1.0"));
        _graphClient.AddDefaultHeader("Authorization",
            InvocationContext.AuthenticationCredentialsProviders.First(p => p.KeyName == "Authorization").Value);
    }

    protected abstract string GetResource();
    
    public async Task SubscribeAsync(IEnumerable<AuthenticationCredentialsProvider> authenticationCredentialsProviders,
        Dictionary<string, string> values)
    {
        var targetSubscription = await GetTargetSubscription();
        var bridgeService = new BridgeService(_bridgeUrl);
        string subscriptionId;
        
        if (targetSubscription is null)
        {
            var createSubscriptionJsonBody = new
            {
                ChangeType = _subscriptionEvent,
                NotificationUrl = _bridgeWebhooksUrl,
                Resource = Resource,
                ExpirationDateTime = (DateTime.Now + TimeSpan.FromMinutes(40000)).ToString("O"),
                ClientState = ApplicationConstants.SharePointClientState
            };
            var subscription =
                await ExecuteRequestAsync<SubscriptionDto>("/subscriptions", Method.Post, createSubscriptionJsonBody);
            subscriptionId = subscription.Id;

            var deltaResult = await ExecuteRequestAsync<ListWrapper<object>>($"{Resource}/delta", Method.Get);
            
            while (deltaResult.ODataNextLink != null)
            {
                var endpoint = deltaResult.ODataNextLink?.Split("v1.0")[1];
                deltaResult = await ExecuteRequestAsync<ListWrapper<object>>(endpoint, Method.Get);
            }
            
            var deltaToken = QueryHelpers.ParseQuery(deltaResult.ODataDeltaLink!.Split("?")[1])["token"];
            await bridgeService.StoreValue(subscriptionId, deltaToken);
        }
        else
            subscriptionId = targetSubscription.Id;
        
        await bridgeService.Subscribe(values["payloadUrl"], subscriptionId, _subscriptionEvent);
    }

    public async Task UnsubscribeAsync(IEnumerable<AuthenticationCredentialsProvider> authenticationCredentialsProviders,
        Dictionary<string, string> values)
    {
        var targetSubscription = await GetTargetSubscription();
        var subscriptionId = targetSubscription.Id;
        
        var bridgeService = new BridgeService(_bridgeUrl);
        var webhooksLeft = await bridgeService.Unsubscribe(values["payloadUrl"], subscriptionId, _subscriptionEvent);

        if (webhooksLeft == 0)
        {
            await bridgeService.DeleteValue(subscriptionId);
            await ExecuteRequestAsync($"/subscriptions/{subscriptionId}", Method.Delete);
        }
    }

    [Period(39995)]
    public async Task RenewSubscription(IEnumerable<AuthenticationCredentialsProvider> authenticationCredentialsProviders,
        Dictionary<string, string> values)
    {
        var targetSubscription = await GetTargetSubscription();
        var updateSubscriptionRequestBody = new
        {
            ExpirationDateTime = (DateTime.Now + TimeSpan.FromMinutes(40000)).ToString("O")
        };
        await ExecuteRequestAsync($"/subscriptions/{targetSubscription.Id}", Method.Patch, updateSubscriptionRequestBody);
    }

    private async Task<SubscriptionDto?> GetTargetSubscription()
    {
        var subscriptions = await ExecuteRequestAsync<SubscriptionWrapper>("/subscriptions", Method.Get);
        var targetSubscription = subscriptions.Value.FirstOrDefault(s => s.Resource == Resource 
                                                                         && s.ChangeType == _subscriptionEvent 
                                                                         && s.NotificationUrl == _bridgeWebhooksUrl);
        return targetSubscription;
    }

    private async Task<RestResponse> ExecuteRequestAsync(string endpoint, Method method, object? jsonBody = null)
    {
        var request = new MicrosoftSharePointRequest(endpoint, method);

        if (jsonBody != null)
            request.AddJsonBody(jsonBody);

        return await _graphClient.ExecuteAsync(request);
    }

    private async Task<T> ExecuteRequestAsync<T>(string endpoint, Method method, object? jsonBody = null)
    {
        var response = await ExecuteRequestAsync(endpoint, method, jsonBody);
        return response.Content.DeserializeObject<T>();
    }
}