using Apps.MicrosoftSharePoint.Dtos;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Invocation;
using Blackbird.Applications.SDK.Extensions.FileManagement.Interfaces;
using Blackbird.Applications.SDK.Extensions.FileManagement.Models.FileDataSourceItems;
using RestSharp;

namespace Apps.MicrosoftSharePoint.DataSourceHandlers
{
    public class FilePickerDataSourceHandler(InvocationContext invocationContext) : BaseInvocable(invocationContext), IAsyncFileDataSourceItemHandler
    {
        private const string RootId = "root";
        private const string RootFolderDisplayName = "My files";

        public async Task<IEnumerable<FileDataItem>> GetFolderContentAsync(FolderContentDataSourceContext context, CancellationToken cancellationToken)
        {
            var client = new SharePointBetaClient(InvocationContext.AuthenticationCredentialsProviders);
            var folderId = string.IsNullOrEmpty(context?.FolderId) ? RootId : context.FolderId!;
            var items = await ListItemsInFolderById(folderId);

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
                        Date = i.CreatedDateTime,
                        IsSelectable = false
                    });
                }
                else
                {
                    result.Add(new Blackbird.Applications.SDK.Extensions.FileManagement.Models.FileDataSourceItems.File
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

        public async Task<IEnumerable<FolderPathItem>> GetFolderPathAsync(FolderPathDataSourceContext context, CancellationToken cancellationToken)
        {
            if (string.IsNullOrEmpty(context?.FileDataItemId))
                return new List<FolderPathItem> { new() { DisplayName = RootFolderDisplayName, Id = RootId } };

            var result = new List<FolderPathItem>();

            try
            {
                var current = await GetFileMetadataById(context.FileDataItemId);
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

        private async Task<List<FileMetadataDto>> ListItemsInFolderById(string folderId)
        {
            var client = new SharePointBetaClient(InvocationContext.AuthenticationCredentialsProviders);
            var items = new List<FileMetadataDto>();
            string? next = folderId == RootId
                    ? "/drive/root/children"
                    : $"/drive/items/{folderId}/children";

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

        private async Task<FileMetadataDto?> GetFileMetadataById(string id)
        {
            var client = new SharePointBetaClient(InvocationContext.AuthenticationCredentialsProviders);
            var request = new SharePointRequest($"/drive/items/{id}", Method.Get, InvocationContext.AuthenticationCredentialsProviders);
            return await client.ExecuteWithHandling<FileMetadataDto>(request);
        }
    }
}
