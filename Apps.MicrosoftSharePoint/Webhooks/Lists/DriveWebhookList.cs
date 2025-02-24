using System.Net;
using Apps.MicrosoftSharePoint.Dtos;
using Apps.MicrosoftSharePoint.Extensions;
using Apps.MicrosoftSharePoint.Models.Responses;
using Apps.MicrosoftSharePoint.Webhooks.Handlers;
using Apps.MicrosoftSharePoint.Webhooks.Inputs;
using Apps.MicrosoftSharePoint.Webhooks.Payload;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Authentication;
using Blackbird.Applications.Sdk.Common.Invocation;
using Blackbird.Applications.Sdk.Common.Webhooks;
using Microsoft.AspNetCore.WebUtilities;
using RestSharp;

namespace Apps.MicrosoftSharePoint.Webhooks.Lists;

[WebhookList]
public class DriveWebhookList : BaseInvocable
{
    private static readonly object LockObject = new();
    
    private readonly IEnumerable<AuthenticationCredentialsProvider> _authenticationCredentialsProviders;

    public DriveWebhookList(InvocationContext invocationContext) : base(invocationContext)
    {
        _authenticationCredentialsProviders = invocationContext.AuthenticationCredentialsProviders;
    }

    [Webhook("On files updated or created", typeof(DriveWebhookHandler), 
        Description = "This webhook is triggered when files are updated or created.")]
    public async Task<WebhookResponse<ListFilesResponse>> OnFilesUpdatedOrCreated(WebhookRequest request, 
        [WebhookParameter] FolderInput folder, [WebhookParameter] ContentTypeInput contentType)
    {
        var payload = DeserializePayload(request);
        var changedFiles = GetChangedItems<FileMetadataDto>(payload.DeltaToken, out var newDeltaToken)
            .Where(item => item.MimeType != null
                           && (folder.ParentFolderId == null || item.ParentReference.Id == folder.ParentFolderId)
                           && (contentType.ContentType == null || item.MimeType == contentType.ContentType))
            .ToList();
        
        if (!changedFiles.Any())
            return new WebhookResponse<ListFilesResponse>
            {
                HttpResponseMessage = new HttpResponseMessage(HttpStatusCode.OK),
                ReceivedWebhookRequestType = WebhookRequestType.Preflight
            };

        await StoreDeltaToken(payload.DeltaToken, newDeltaToken);
        return new WebhookResponse<ListFilesResponse>
        {
            HttpResponseMessage = new HttpResponseMessage(HttpStatusCode.OK),
            Result = new ListFilesResponse { Files = changedFiles }
        };
    }
    
    [Webhook("On folders updated or created", typeof(DriveWebhookHandler), 
        Description = "This webhook is triggered when folders are updated or created.")]
    public async Task<WebhookResponse<ListFoldersResponse>> OnFoldersUpdatedOrCreated(WebhookRequest request, 
        [WebhookParameter] FolderInput folder)
    {
        var payload = DeserializePayload(request);
        var changedFolders = GetChangedItems<FolderMetadataDto>(payload.DeltaToken, out var newDeltaToken)
            .Where(item => item.ChildCount != null 
                           && item.ParentReference!.Id != null  
                           && (folder.ParentFolderId == null || item.ParentReference.Id == folder.ParentFolderId))
            .ToList();
        
        if (!changedFolders.Any())
            return new WebhookResponse<ListFoldersResponse>
            {
                HttpResponseMessage = new HttpResponseMessage(HttpStatusCode.OK),
                ReceivedWebhookRequestType = WebhookRequestType.Preflight
            };

        await StoreDeltaToken(payload.DeltaToken, newDeltaToken);
        return new WebhookResponse<ListFoldersResponse>
        {
            HttpResponseMessage = new HttpResponseMessage(HttpStatusCode.OK),
            Result = new ListFoldersResponse { Folders = changedFolders }
        };
    }

    private List<T> GetChangedItems<T>(string deltaToken, out string newDeltaToken)
    {
        var client = new SharePointBetaClient(_authenticationCredentialsProviders);
        var items = new List<T>();
        var request = new SharePointRequest($"/drive/root/delta?token={deltaToken}", Method.Get, 
            _authenticationCredentialsProviders);
        var result = client.ExecuteWithHandling<ListWrapper<T>>(request).Result;
        items.AddRange(result.Value);

        while (result.ODataNextLink != null)
        {
            var endpoint = result.ODataNextLink?.Split("v1.0")[1];
            request = new SharePointRequest(endpoint, Method.Get, _authenticationCredentialsProviders);
            result = client.ExecuteWithHandling<ListWrapper<T>>(request).Result;
            items.AddRange(result.Value);
        }
        
        newDeltaToken = QueryHelpers.ParseQuery(result.ODataDeltaLink!.Split("?")[1])["token"];
        return items;
    }

    private EventPayload DeserializePayload(WebhookRequest request) 
        => request.Body.ToString().DeserializeObject<EventPayload>();

    private async Task StoreDeltaToken(string oldDeltaToken, string newDeltaToken)
    {
        string bridgeWebhooksUrl = InvocationContext.UriInfo.BridgeServiceUrl.ToString().TrimEnd('/') + $"/webhooks/{ApplicationConstants.AppName}";
        
        var siteId = InvocationContext.AuthenticationCredentialsProviders.First(p => p.KeyName == "SiteId").Value;
        var resource = $"/sites/{siteId}/drive/root";
        var sharePointClient = new SharePointClient();
        var subscriptionsRequest = new SharePointRequest("/subscriptions", Method.Get, _authenticationCredentialsProviders);
        var response = await sharePointClient.ExecuteAsync(subscriptionsRequest);
        var subscriptions = response.Content.DeserializeObject<SubscriptionWrapper>().Value;
        var targetSubscription = subscriptions.Single(s => s.Resource == resource
                                                                   && s.NotificationUrl == bridgeWebhooksUrl);

        var bridgeService = new BridgeService(InvocationContext.UriInfo.BridgeServiceUrl.ToString().TrimEnd('/'));
        
        lock (LockObject)
        {
            var storedDeltaToken = bridgeService.RetrieveValue(targetSubscription.Id).Result.Trim('"');
            if (storedDeltaToken == oldDeltaToken)
                bridgeService.StoreValue(targetSubscription.Id, newDeltaToken).Wait();
        }
    }
}