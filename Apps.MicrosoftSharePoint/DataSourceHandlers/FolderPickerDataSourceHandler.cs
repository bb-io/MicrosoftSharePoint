using RestSharp;
using Apps.MicrosoftSharePoint.Dtos;
using Apps.MicrosoftSharePoint.Models.Identifiers;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Invocation;
using Blackbird.Applications.SDK.Extensions.FileManagement.Interfaces;
using Blackbird.Applications.SDK.Extensions.FileManagement.Models.FileDataSourceItems;

namespace Apps.MicrosoftSharePoint.DataSourceHandlers;

public class FolderPickerDataSourceHandler(
    InvocationContext invocationContext,
    [ActionParameter] FolderIdentifier folder) 
    : BaseInvocable(invocationContext), IAsyncFileDataSourceItemHandler
{
    private const string RootId = "root";
    private const string RootFolderDisplayName = "My files";

    public async Task<IEnumerable<FileDataItem>> GetFolderContentAsync(FolderContentDataSourceContext context,CancellationToken cancellationToken)
    {
        var client = new SharePointBetaClient(InvocationContext.AuthenticationCredentialsProviders);
        var folderId = string.IsNullOrEmpty(context?.FolderId) ? RootId : context.FolderId!;
        var items = await ListItemsInFolderById(folderId, cancellationToken);

        return items
            .Where(x => string.IsNullOrEmpty(x.MimeType))
            .Select(x => new Folder
            {
                Id = x.FileId,
                DisplayName = x.Name,
                Date = x.CreatedDateTime,
                IsSelectable = true
            })
            .Cast<FileDataItem>()
            .ToList();
    }

    public async Task<IEnumerable<FolderPathItem>> GetFolderPathAsync(FolderPathDataSourceContext context,CancellationToken cancellationToken)
    {
        var client = new SharePointBetaClient(InvocationContext.AuthenticationCredentialsProviders);
        if (string.IsNullOrEmpty(context?.FileDataItemId))
            return new List<FolderPathItem>
            {
                new() { DisplayName = RootFolderDisplayName, Id = RootId }
            };

        var result = new List<FolderPathItem>();

        try
        {
            var current = await GetFileMetadataById(context.FileDataItemId!);
            var parentId = current?.ParentReference?.Id;

            while (!string.IsNullOrEmpty(parentId))
            {
                var parent = await GetFileMetadataById(parentId!);
                if (parent == null) break;

                result = result
                    .Prepend(new FolderPathItem
                    {
                        DisplayName = parent.Name,
                        Id = parent.FileId
                    })
                    .ToList();

                parentId = parent.ParentReference?.Id;
            }

            var root = result.FirstOrDefault();
            if (root != null)
            {
                root.DisplayName = RootFolderDisplayName;
                root.Id = RootId;
            }
            else
            {
                result.Add(new FolderPathItem { DisplayName = RootFolderDisplayName, Id = RootId });
            }
        }
        catch
        {
            result.Clear();
            result.Add(new FolderPathItem { DisplayName = RootFolderDisplayName, Id = RootId });
        }

        return result;
    }

    private async Task<List<FileMetadataDto>> ListItemsInFolderById(string folderId, CancellationToken ct)
    {
        var client = new SharePointBetaClient(InvocationContext.AuthenticationCredentialsProviders);
        var items = new List<FileMetadataDto>();

        var drivePrefix = string.IsNullOrEmpty(folder.DriveId)
            ? "/drive"
            : $"/drives/{folder.DriveId}";
        string? next = folderId == RootId
            ? $"{drivePrefix}/root/children"
            : $"{drivePrefix}/items/{folderId}/children";

        do
        {
            var request = Uri.IsWellFormedUriString(next, UriKind.Absolute)
                ? new SharePointRequest(new Uri(next!).ToString(), Method.Get, InvocationContext.AuthenticationCredentialsProviders)
                : new SharePointRequest(next!, Method.Get, InvocationContext.AuthenticationCredentialsProviders);

            var page = await client.ExecuteWithHandling<ListWrapper<FileMetadataDto>>(request);
            items.AddRange(page?.Value ?? Array.Empty<FileMetadataDto>());
            next = page?.ODataNextLink;
        }
        while (!string.IsNullOrEmpty(next));

        return items;
    }

    private async Task<FileMetadataDto?> GetFileMetadataById(string id)
    {
        var client = new SharePointBetaClient(InvocationContext.AuthenticationCredentialsProviders);
        var request = new SharePointRequest($"/drive/items/{id}", Method.Get, InvocationContext.AuthenticationCredentialsProviders);
        return await client.ExecuteWithHandling<FileMetadataDto>(request);
    }
}
