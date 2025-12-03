using System.Net;
using Apps.MicrosoftSharePoint.Dtos;
using Apps.MicrosoftSharePoint.Extensions;
using Apps.MicrosoftSharePoint.Helper;
using Apps.MicrosoftSharePoint.Models.Entities;
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
        DriveEntity defaultDrive = await GetDefaultDrive();

        var allowedParentIds = await GetAllowedParentIds(location, defaultDrive);

        var changedFiles = GetChangedItems<FileMetadataDto>(payload.DeltaToken, location, out var newDeltaToken)
            .Where(item => item.MimeType != null
                           && (folder.FolderId == null || (item.ParentReference?.Id != null && allowedParentIds.Contains(item.ParentReference.Id)))
                           && (contentType.ContentType == null || item.MimeType == contentType.ContentType))
            .ToList();

        if (!changedFiles.Any())
            return new WebhookResponse<ListFilesResponse>
            {
                HttpResponseMessage = new HttpResponseMessage(HttpStatusCode.OK),
                ReceivedWebhookRequestType = WebhookRequestType.Preflight
            };

        foreach (var file in changedFiles)
        {
            var currentDriveId = location.DriveId ?? defaultDrive!.Id;
            file.FileId = ItemIdParser.Format(currentDriveId, file.FileId, defaultDrive!.Id); 
            if (file.ParentReference != null && !string.IsNullOrEmpty(file.ParentReference.Id))
                file.ParentReference.Id = ItemIdParser.Format(currentDriveId, file.ParentReference.Id, defaultDrive!.Id);
        }

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
        DriveEntity defaultDrive = await GetDefaultDrive();

        var allowedParentIds = await GetAllowedParentIds(location, defaultDrive);

        var changedFolders = GetChangedItems<FolderMetadataDto>(payload.DeltaToken, location, out var newDeltaToken)
            .Where(item => item.ChildCount != null
                           && item.ParentReference!.Id != null
                           && (folder.FolderId == null || allowedParentIds.Contains(item.ParentReference.Id)))
            .ToList();

        if (!changedFolders.Any())
            return new WebhookResponse<ListFoldersResponse>
            {
                HttpResponseMessage = new HttpResponseMessage(HttpStatusCode.OK),
                ReceivedWebhookRequestType = WebhookRequestType.Preflight
            };

        foreach (var changedfolder in changedFolders)
        {
            var currentDriveId = location.DriveId ?? defaultDrive!.Id;
            changedfolder.Id = ItemIdParser.Format(currentDriveId, changedfolder.Id, defaultDrive!.Id);
            if (changedfolder.ParentReference != null && !string.IsNullOrEmpty(changedfolder.ParentReference.Id))
                changedfolder.ParentReference.Id = ItemIdParser.Format(currentDriveId, changedfolder.ParentReference.Id, defaultDrive!.Id);
        }

        await StoreDeltaToken(payload.DeltaToken, newDeltaToken, location);
        return new WebhookResponse<ListFoldersResponse>
        {
            HttpResponseMessage = new HttpResponseMessage(HttpStatusCode.OK),
            Result = new ListFoldersResponse { Folders = changedFolders }
        };
    }

    private async Task<List<string>> GetAllowedParentIds(ItemLocationDto location, DriveEntity defaultDrive)
    {
        var allowedParentIds = new List<string>();
        if (!string.Equals(location.ItemId, "root", StringComparison.OrdinalIgnoreCase))
            allowedParentIds.Add(location.ItemId);
        else
        {
            var rootFolder = await GetRootFolder(location);
            if (rootFolder?.Id != null)
                allowedParentIds.Add(rootFolder.Id);

            if (location.IsDefaultDrive)
            {
                if (defaultDrive?.Id != null)
                    allowedParentIds.Add(defaultDrive.Id);
            }
            else
                allowedParentIds.Add(location.DriveId!);
        }
        return allowedParentIds;
    }

    private async Task<FolderMetadataDto> GetRootFolder(ItemLocationDto location)
    {
        var client = new SharePointBetaClient(creds);
        string endpoint;

        if (location.IsDefaultDrive)
            endpoint = "/drive/root";
        else
            endpoint = $"/drives/{location.DriveId}/root";

        var request = new SharePointRequest(endpoint, Method.Get, creds);
        return await client.ExecuteWithHandling<FolderMetadataDto>(request);
    }

    private async Task<DriveEntity> GetDefaultDrive()
    {
        var creds = InvocationContext.AuthenticationCredentialsProviders;
        var siteId = creds.First(x => x.KeyName == "SiteId").Value;
        var client = new SharePointBetaClient(creds);
        var request = new SharePointRequest($"/sites/{siteId}/drive", Method.Get, creds);
        return await client.ExecuteWithHandling<DriveEntity>(request);
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