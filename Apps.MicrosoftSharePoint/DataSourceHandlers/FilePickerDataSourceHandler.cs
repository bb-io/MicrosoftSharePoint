using RestSharp;
using Apps.MicrosoftSharePoint.Dtos;
using Apps.MicrosoftSharePoint.Models.Entities;
using Apps.MicrosoftSharePoint.Models.Responses;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Invocation;
using Blackbird.Applications.SDK.Extensions.FileManagement.Interfaces;
using Blackbird.Applications.SDK.Extensions.FileManagement.Models.FileDataSourceItems;
using File = Blackbird.Applications.SDK.Extensions.FileManagement.Models.FileDataSourceItems.File;

namespace Apps.MicrosoftSharePoint.DataSourceHandlers;

public class FilePickerDataSourceHandler(InvocationContext invocationContext)
    : BaseInvocable(invocationContext), IAsyncFileDataSourceItemHandler
{
    public async Task<IEnumerable<FileDataItem>> GetFolderContentAsync(FolderContentDataSourceContext context, CancellationToken cancellationToken)
    {
        if (string.IsNullOrEmpty(context.FolderId))
            return await GetDrivesListAsFolders();

        bool isDefaultDrive = !context.FolderId.Contains('#');
        string? driveId = null;
        string? itemId = null;

        if (isDefaultDrive)
            itemId = context.FolderId;
        else
        {
            var parts = context.FolderId.Split('#');
            driveId = parts[0];
            itemId = parts.Length > 1 ? parts[1] : "root";
        }

        var items = await ListItems(isDefaultDrive, driveId, itemId);

        var result = new List<FileDataItem>();
        foreach (var i in items)
        {
            var isFolder = string.IsNullOrEmpty(i.MimeType);

            var idForUi = isDefaultDrive ? i.FileId : $"{driveId}#{i.FileId}";

            if (isFolder)
            {
                result.Add(new Folder
                {
                    Id = idForUi,
                    DisplayName = i.Name,
                    Date = i.LastModifiedDateTime,
                    IsSelectable = false
                });
            }
            else
            {
                result.Add(new File
                {
                    Id = idForUi,
                    DisplayName = i.Name,
                    Date = i.LastModifiedDateTime,
                    Size = i.Size,
                    IsSelectable = true
                });
            }
        }

        return result;
    }

    public async Task<IEnumerable<FolderPathItem>> GetFolderPathAsync(FolderPathDataSourceContext context, CancellationToken cancellationToken)
    {
        if (string.IsNullOrEmpty(context.FileDataItemId))
            return [];

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
            var firstItem = await GetItemInDrive(driveId, currentItemId);

            if (firstItem != null)
            {
                bool isFile = Path.HasExtension(firstItem.Name);
                if (!isFile)
                {
                    path.Add(new FolderPathItem
                    {
                        Id = $"{driveId}#{firstItem.FileId}",
                        DisplayName = firstItem.Name
                    });
                }

                currentItemId = firstItem.ParentReference?.Id;

                while (!string.IsNullOrEmpty(currentItemId))
                {
                    var item = await GetItemInDrive(driveId, currentItemId);
                    if (item == null) break;

                    bool isRootFolder = item.Name.Equals("root", StringComparison.OrdinalIgnoreCase) 
                        || item.ParentReference == null;

                    if (!isRootFolder)
                    {
                        path.Add(new FolderPathItem
                        {
                            Id = $"{driveId}#{item.FileId}",
                            DisplayName = item.Name
                        });
                    }

                    currentItemId = item.ParentReference?.Id;
                }
            }

            var drive = await GetDriveById(driveId);
            if (drive != null)
                path.Add(new FolderPathItem { DisplayName = drive.Name, Id = driveId });
        }

        path.Add(new FolderPathItem { DisplayName = "Home", Id = string.Empty });

        path.Reverse();
        return path;
    }

    private async Task<IEnumerable<FileDataItem>> GetDrivesListAsFolders()
    {
        var allDrivesTask = GetDrives();
        var defaultDriveTask = GetDefaultDrive();

        await Task.WhenAll(allDrivesTask, defaultDriveTask);

        var drivesResponse = allDrivesTask.Result;
        var defaultDrive = defaultDriveTask.Result;

        var folders = new List<FileDataItem>();

        foreach (var drive in drivesResponse.Value)
        {
            string idForNavigation;

            if (drive.Id == defaultDrive.Id)
                idForNavigation = "root";
            else
                idForNavigation = $"{drive.Id}#root";

            folders.Add(new Folder
            {
                Id = idForNavigation,
                DisplayName = drive.Name,
                IsSelectable = false,
                Date = drive.LastModified
            });
        }

        return folders;
    }

    private async Task<DriveEntity> GetDefaultDrive()
    {
        var creds = InvocationContext.AuthenticationCredentialsProviders;
        var siteId = creds.First(x => x.KeyName == "SiteId").Value;

        var client = new SharePointBetaClient(creds);
        var request = new SharePointRequest($"/sites/{siteId}/drive", Method.Get, creds);
        return await client.ExecuteWithHandling<DriveEntity>(request);
    }

    private async Task<List<FileMetadataDto>> ListItems(bool isDefaultDrive, string? driveId, string itemId)
    {
        var client = new SharePointBetaClient(InvocationContext.AuthenticationCredentialsProviders);
        string endpoint;

        if (isDefaultDrive)
        {
            endpoint = itemId.Equals("root", StringComparison.OrdinalIgnoreCase)
                ? "/drive/root/children"
                : $"/drive/items/{itemId}/children";
        }
        else
        {
            endpoint = itemId.Equals("root", StringComparison.OrdinalIgnoreCase)
                ? $"/drives/{driveId}/root/children"
                : $"/drives/{driveId}/items/{itemId}/children";
        }

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
        var creds = InvocationContext.AuthenticationCredentialsProviders;
        var siteId = creds.First(x => x.KeyName == "SiteId").Value;

        var client = new SharePointClient();
        var request = new SharePointRequest($"/sites/{siteId}/drives", Method.Get, creds);
        return await client.ExecuteWithHandling<ListResponse<DriveEntity>>(request);
    }
}