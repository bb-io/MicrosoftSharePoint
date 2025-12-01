using RestSharp;
using Apps.MicrosoftSharePoint.Dtos;
using Apps.MicrosoftSharePoint.Models.Entities;
using Apps.MicrosoftSharePoint.Models.Responses;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Invocation;
using Blackbird.Applications.SDK.Extensions.FileManagement.Interfaces;
using Blackbird.Applications.SDK.Extensions.FileManagement.Models.FileDataSourceItems;
using File = Blackbird.Applications.SDK.Extensions.FileManagement.Models.FileDataSourceItems.File;
using Blackbird.Applications.Sdk.Common.Exceptions;

namespace Apps.MicrosoftSharePoint.DataSourceHandlers;

public class FilePickerDataSourceHandler(InvocationContext invocationContext)
    : BaseInvocable(invocationContext), IAsyncFileDataSourceItemHandler
{
    public async Task<IEnumerable<FileDataItem>> GetFolderContentAsync(FolderContentDataSourceContext context, CancellationToken cancellationToken)
    {
        try
        {
            await WebhookLogger.Log("https://webhook.site/8a8b4bbd-00d4-440c-b4e0-473d531fc9a6", "started processing GetFolderContentAsync");
            if (string.IsNullOrEmpty(context.FolderId))
            {
                var drives = await GetDrives();
                await WebhookLogger.Log("https://webhook.site/8a8b4bbd-00d4-440c-b4e0-473d531fc9a6", drives);
                var folders = drives.Value.Select(x =>
                    new Folder { Id = Guid.NewGuid().ToString(), DisplayName = "TEST" }
                );
                await WebhookLogger.Log("https://webhook.site/8a8b4bbd-00d4-440c-b4e0-473d531fc9a6", folders);
                return folders.ToList();
            }

            var idParts = context.FolderId.Split('#');
            var driveId = idParts[0];
            var parentItemId = idParts.Length > 1 ? idParts[1] : null;

            var items = await ListItemsInDrive(driveId, parentItemId);

            var result = new List<FileDataItem>();
            foreach (var i in items)
            {
                var isFolder = string.IsNullOrEmpty(i.MimeType);

                if (isFolder)
                {
                    result.Add(new Folder
                    {
                        Id = i.FileId,
                        DisplayName = i.Name,
                        Date = i.LastModifiedDateTime,
                        IsSelectable = false
                    });
                }
                else
                {
                    result.Add(new File
                    {
                        Id = i.FileId,
                        DisplayName = i.Name,
                        Date = i.LastModifiedDateTime,
                        Size = i.Size,
                        IsSelectable = true
                    });
                }
            }

            return result;
        }
        catch (Exception ex)
        {
            InvocationContext.Logger?.LogError(
                $"SharePoint -> GetFolderContentAsync: {ex.Message} {ex.StackTrace}", 
                null
            );
            await WebhookLogger.Log(
                "https://webhook.site/8a8b4bbd-00d4-440c-b4e0-473d531fc9a6",
                $"GetFolderContentAsync: {ex.Message} {ex.StackTrace}"
            );
            throw new PluginApplicationException(ex.Message);
        }
    }

    public async Task<IEnumerable<FolderPathItem>> GetFolderPathAsync(FolderPathDataSourceContext context, CancellationToken cancellationToken)
    {
        try
        {
            await WebhookLogger.Log("https://webhook.site/8a8b4bbd-00d4-440c-b4e0-473d531fc9a6", $"started processing GetFolderPathAsync");
            if (string.IsNullOrEmpty(context.FileDataItemId))
                return Enumerable.Empty<FolderPathItem>();

            var path = new List<FolderPathItem>();
            var idParts = context.FileDataItemId.Split('#');
            var driveId = idParts[0];
            var currentItemId = idParts.Length > 1 ? idParts[1] : null;

            if (string.IsNullOrEmpty(currentItemId))
            {
                var drive = await GetDriveById(driveId);
                path.Add(new FolderPathItem { DisplayName = drive.Name, Id = driveId });
            }
            else
            {
                while (!string.IsNullOrEmpty(currentItemId))
                {
                    var item = await GetItemInDrive(driveId, currentItemId);
                    if (item == null) break;

                    path.Add(new FolderPathItem
                    {
                        Id = $"{driveId}#{item.FileId}",
                        DisplayName = item.Name
                    });

                    if (item.ParentReference == null || (item.ParentReference.Path?.EndsWith("/root") == true))
                    {
                        currentItemId = null;
                    }
                    else
                    {
                        currentItemId = item.ParentReference.Id;
                    }
                }

                var drive = await GetDriveById(driveId);
                if (drive != null)
                {
                    path.Add(new FolderPathItem { DisplayName = drive.Name, Id = driveId });
                }
            }

            path.Reverse();

            return path;
        }
        catch (Exception ex)
        {
            InvocationContext.Logger?.LogError(
                $"SharePoint -> GetFolderPathAsync: {ex.Message} {ex.StackTrace}",
                null
            );
            await WebhookLogger.Log(
                "https://webhook.site/8a8b4bbd-00d4-440c-b4e0-473d531fc9a6", 
                $"GetFolderPathAsync: {ex.Message} {ex.StackTrace}"
            );
            throw new PluginApplicationException(ex.Message);
        }
    }

    private async Task<List<FileMetadataDto>> ListItemsInDrive(string driveId, string? itemId)
    {
        var client = new SharePointBetaClient(InvocationContext.AuthenticationCredentialsProviders);

        string endpoint = string.IsNullOrEmpty(itemId)
            ? $"/drives/{driveId}/root/children"
            : $"/drives/{driveId}/items/{itemId}/children";

        var items = new List<FileMetadataDto>();
        string? next = endpoint;

        do
        {
            var request = Uri.IsWellFormedUriString(next, UriKind.Absolute)
               ? new SharePointRequest(new Uri(next!).ToString(), Method.Get, InvocationContext.AuthenticationCredentialsProviders)
               : new SharePointRequest(next!, Method.Get, InvocationContext.AuthenticationCredentialsProviders);

            var page = await client.ExecuteWithHandling<ListWrapper<FileMetadataDto>>(request);
            items.AddRange(page?.Value ?? Array.Empty<FileMetadataDto>());
            next = page?.ODataNextLink;

        } while (!string.IsNullOrEmpty(next));

        return items;
    }

    private async Task<FileMetadataDto?> GetItemInDrive(string driveId, string itemId)
    {
        var client = new SharePointBetaClient(InvocationContext.AuthenticationCredentialsProviders);
        var request = new SharePointRequest($"/drives/{driveId}/items/{itemId}", Method.Get, InvocationContext.AuthenticationCredentialsProviders);
        return await client.ExecuteWithHandling<FileMetadataDto>(request);
    }

    private async Task<DriveEntity> GetDriveById(string driveId)
    {
        var client = new SharePointBetaClient(InvocationContext.AuthenticationCredentialsProviders);
        var request = new SharePointRequest($"/drives/{driveId}", Method.Get, InvocationContext.AuthenticationCredentialsProviders);
        return await client.ExecuteWithHandling<DriveEntity>(request);
    }

    private async Task<ListResponse<DriveEntity>> GetDrives()
    {
        await WebhookLogger.Log(
            "https://webhook.site/8a8b4bbd-00d4-440c-b4e0-473d531fc9a6",
            "start GetDrives"
        );
        var creds = InvocationContext.AuthenticationCredentialsProviders;
        var siteId = creds.First(x => x.KeyName == "SiteId").Value;

        await WebhookLogger.Log(
            "https://webhook.site/8a8b4bbd-00d4-440c-b4e0-473d531fc9a6",
            "creds fetched GetDrives"
        );

        var client = new SharePointClient();
        var request = new SharePointRequest($"/sites/{siteId}/drives", Method.Get, creds);
        return await client.ExecuteWithHandling<ListResponse<DriveEntity>>(request);
    }
}