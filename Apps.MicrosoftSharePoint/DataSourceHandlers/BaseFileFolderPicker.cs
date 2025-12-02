using RestSharp;
using Apps.MicrosoftSharePoint.Dtos;
using Apps.MicrosoftSharePoint.Helper;
using Apps.MicrosoftSharePoint.Models.Entities;
using Apps.MicrosoftSharePoint.Models.Responses;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Invocation;
using Blackbird.Applications.SDK.Extensions.FileManagement.Models.FileDataSourceItems;
using File = Blackbird.Applications.SDK.Extensions.FileManagement.Models.FileDataSourceItems.File;

namespace Apps.MicrosoftSharePoint.DataSourceHandlers;

public class BaseFileFolderPicker(InvocationContext invocationContext) : BaseInvocable(invocationContext)
{
    public async Task<IEnumerable<FileDataItem>> GetFolderContent(
        string? folderId,
        bool filesAreSelectable,
        bool foldersAreSelectable)
    {
        if (string.IsNullOrEmpty(folderId))
            return await GetDrivesListAsFolders();

        var location = ItemIdParser.Parse(folderId);

        var defaultDrive = await GetDefaultDrive();

        var items = await ListItems(location);
        var result = new List<FileDataItem>();
        foreach (var i in items)
        {
            var isFolder = string.IsNullOrEmpty(i.MimeType);

            var actualDriveId = location.DriveId ?? defaultDrive.Id;
            var generatedId = ItemIdParser.Format(actualDriveId, i.FileId, defaultDrive.Id);

            if (isFolder)
            {
                result.Add(new Folder
                {
                    Id = generatedId,
                    DisplayName = i.Name,
                    Date = i.LastModifiedDateTime,
                    IsSelectable = foldersAreSelectable
                });
            }
            else
            {
                result.Add(new File
                {
                    Id = generatedId,
                    DisplayName = i.Name,
                    Date = i.LastModifiedDateTime,
                    Size = i.Size,
                    IsSelectable = filesAreSelectable
                });
            }
        }

        return result;
    }

    public async Task<IEnumerable<FolderPathItem>> GetFolderPath(string? fileDataItemId)
    {
        if (string.IsNullOrEmpty(fileDataItemId))
            return [];

        var location = ItemIdParser.Parse(fileDataItemId);

        var defaultDrive = await GetDefaultDrive();
        if (location.IsDefaultDrive)
            location.DriveId = defaultDrive.Id;
        
        var driveId = location.DriveId!;
        var currentItemId = location.ItemId;

        var path = new List<FolderPathItem>();
        if (!string.IsNullOrEmpty(currentItemId) && !currentItemId.Equals("root", StringComparison.OrdinalIgnoreCase))
        {
            var firstItem = await GetItemInDrive(driveId, currentItemId);

            if (firstItem != null)
            {
                bool isFile = !string.IsNullOrEmpty(firstItem.MimeType);

                if (!isFile)
                {
                    path.Add(new FolderPathItem
                    {
                        Id = ItemIdParser.Format(driveId, firstItem.FileId, defaultDrive.Id),
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

                    if (isRootFolder) break;

                    path.Add(new FolderPathItem
                    {
                        Id = ItemIdParser.Format(driveId, item.FileId, defaultDrive.Id),
                        DisplayName = item.Name
                    });

                    currentItemId = item.ParentReference?.Id;
                }
            }
        }

        var drive = await GetDriveById(driveId);
        if (drive != null)
        {
            string driveNodeId = ItemIdParser.Format(driveId, "root", defaultDrive.Id);
            path.Add(new FolderPathItem { DisplayName = drive.Name, Id = driveNodeId });
        }

        path.Add(new FolderPathItem { DisplayName = "Home", Id = string.Empty });

        path.Reverse();
        return path;
    }

    private async Task<List<FileMetadataDto>> ListItems(ItemLocationDto location)
    {
        var client = new SharePointBetaClient(InvocationContext.AuthenticationCredentialsProviders);
        string endpoint;

        if (location.IsDefaultDrive)
        {
            endpoint = location.ItemId.Equals("root", StringComparison.OrdinalIgnoreCase)
                ? "/drive/root/children"
                : $"/drive/items/{location.ItemId}/children";
        }
        else
        {
            endpoint = location.ItemId.Equals("root", StringComparison.OrdinalIgnoreCase)
                ? $"/drives/{location.DriveId}/root/children"
                : $"/drives/{location.DriveId}/items/{location.ItemId}/children";
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
            string idForNavigation = ItemIdParser.Format(drive.Id, "root", defaultDrive.Id);
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

    private async Task<DriveEntity> GetDefaultDrive()
    {
        var creds = InvocationContext.AuthenticationCredentialsProviders;
        var siteId = creds.First(x => x.KeyName == "SiteId").Value;

        var client = new SharePointBetaClient(creds);
        var request = new SharePointRequest($"/sites/{siteId}/drive", Method.Get, creds);
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
