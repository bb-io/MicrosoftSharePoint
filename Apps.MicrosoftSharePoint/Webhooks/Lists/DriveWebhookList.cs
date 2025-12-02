using System.Net;
using Apps.MicrosoftSharePoint.Dtos;
using Apps.MicrosoftSharePoint.Extensions;
using Apps.MicrosoftSharePoint.Helper;
using Apps.MicrosoftSharePoint.Models.Identifiers;
using Apps.MicrosoftSharePoint.Models.Responses;
using Apps.MicrosoftSharePoint.Webhooks.Handlers;
using Apps.MicrosoftSharePoint.Webhooks.Inputs;
using Apps.MicrosoftSharePoint.Webhooks.Payload;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Authentication;
using Blackbird.Applications.Sdk.Common.Invocation;
using Blackbird.Applications.Sdk.Common.Webhooks;
using Blackbird.Applications.SDK.Blueprints;
using Microsoft.AspNetCore.WebUtilities;
using RestSharp;

namespace Apps.MicrosoftSharePoint.Webhooks.Lists;

[WebhookList]
public class DriveWebhookList(InvocationContext invocationContext) : BaseInvocable(invocationContext)
{
    private static readonly object LockObject = new();

    private readonly IEnumerable<AuthenticationCredentialsProvider> creds =
        invocationContext.AuthenticationCredentialsProviders;

    [BlueprintEventDefinition(BlueprintEvent.FilesCreatedOrUpdated)]
    [Webhook("On files updated or created", typeof(DriveWebhookHandler),
        Description = "This webhook is triggered when files are updated or created.")]
    public async Task<WebhookResponse<ListFilesResponse>> OnFilesUpdatedOrCreated(
        WebhookRequest request,
        [WebhookParameter] FolderIdentifier folder, 
        [WebhookParameter] ContentTypeInput contentType)
    {
        var payload = DeserializePayload(request);

        var location = ItemIdParser.Parse(folder.FolderId);

        var changedFiles = GetChangedItems<FileMetadataDto>(payload.DeltaToken, location, out var newDeltaToken)
            .Where(item => item.MimeType != null
                           && (folder.FolderId == null || item.ParentReference.Id == location.ItemId)
                           && (contentType.ContentType == null || item.MimeType == contentType.ContentType))
            .ToList();

        if (!changedFiles.Any())
            return new WebhookResponse<ListFilesResponse>
            {
                HttpResponseMessage = new HttpResponseMessage(HttpStatusCode.OK),
                ReceivedWebhookRequestType = WebhookRequestType.Preflight
            };

        await StoreDeltaToken(payload.DeltaToken, newDeltaToken, location);
        return new WebhookResponse<ListFilesResponse>
        {
            HttpResponseMessage = new HttpResponseMessage(HttpStatusCode.OK),
            Result = new ListFilesResponse { Files = changedFiles }
        };
    }

    [Webhook("On folders updated or created", typeof(DriveWebhookHandler),
        Description = "This webhook is triggered when folders are updated or created.")]
    public async Task<WebhookResponse<ListFoldersResponse>> OnFoldersUpdatedOrCreated(
        WebhookRequest request,
        [WebhookParameter] FolderIdentifier folder)
    {
        var payload = DeserializePayload(request);

        var location = ItemIdParser.Parse(folder.FolderId);

        var changedFolders = GetChangedItems<FolderMetadataDto>(payload.DeltaToken, location, out var newDeltaToken)
            .Where(item => item.ChildCount != null
                           && item.ParentReference!.Id != null
                           && (folder.FolderId == null || item.ParentReference.Id == location.ItemId))
            .ToList();

        if (!changedFolders.Any())
            return new WebhookResponse<ListFoldersResponse>
            {
                HttpResponseMessage = new HttpResponseMessage(HttpStatusCode.OK),
                ReceivedWebhookRequestType = WebhookRequestType.Preflight
            };

        await StoreDeltaToken(payload.DeltaToken, newDeltaToken, location);
        return new WebhookResponse<ListFoldersResponse>
        {
            HttpResponseMessage = new HttpResponseMessage(HttpStatusCode.OK),
            Result = new ListFoldersResponse { Folders = changedFolders }
        };
    }

    private List<T> GetChangedItems<T>(string deltaToken, ItemLocationDto location, out string newDeltaToken)
    {
        var client = new SharePointBetaClient(creds);
        var items = new List<T>();

        string baseEndpoint;
        if (location.IsDefaultDrive)
            baseEndpoint = "/drive/root";
        else
            baseEndpoint = $"/drives/{location.DriveId}/root";

        var request = new SharePointRequest($"{baseEndpoint}/delta?token={deltaToken}", Method.Get, creds);

        var result = client.ExecuteWithHandling<ListWrapper<T>>(request).Result;
        items.AddRange(result.Value);

        while (result.ODataNextLink != null)
        {
            var nextLink = result.ODataNextLink;

            request = Uri.IsWellFormedUriString(nextLink, UriKind.Absolute)
                ? new SharePointRequest(new Uri(nextLink).ToString(), Method.Get, creds)
                : new SharePointRequest(nextLink, Method.Get, creds);

            result = client.ExecuteWithHandling<ListWrapper<T>>(request).Result;
            items.AddRange(result.Value);
        }

        newDeltaToken = QueryHelpers.ParseQuery(result.ODataDeltaLink!.Split("?")[1])["token"];
        return items;
    }

    private EventPayload DeserializePayload(WebhookRequest request)
        => request.Body.ToString().DeserializeObject<EventPayload>();

    private async Task StoreDeltaToken(string oldDeltaToken, string newDeltaToken, ItemLocationDto location)
    {
        string bridgeWebhooksUrl = InvocationContext.UriInfo.BridgeServiceUrl.ToString().TrimEnd('/') + $"/webhooks/{ApplicationConstants.AppName}";

        var siteId = InvocationContext.AuthenticationCredentialsProviders.First(p => p.KeyName == "SiteId").Value;

        string resource;
        if (location.IsDefaultDrive)
            resource = $"/sites/{siteId}/drive/root";
        else
            resource = $"/sites/{siteId}/drives/{location.DriveId}/root";

        var sharePointClient = new SharePointClient();
        var subscriptionsRequest = new SharePointRequest("/subscriptions", Method.Get, creds);
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