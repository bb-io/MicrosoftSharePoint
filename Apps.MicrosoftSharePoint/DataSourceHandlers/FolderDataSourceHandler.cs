using Apps.MicrosoftSharePoint.Dtos;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Dynamic;
using Blackbird.Applications.Sdk.Common.Invocation;
using RestSharp;

namespace Apps.MicrosoftSharePoint.DataSourceHandlers;

public class FolderDataSourceHandler : BaseInvocable, IAsyncDataSourceHandler
{
    public FolderDataSourceHandler(InvocationContext invocationContext) : base(invocationContext)
    {
    }

    public async Task<Dictionary<string, string>> GetDataAsync(DataSourceContext context, CancellationToken cancellationToken)
    {
        var client = new SharePointBetaClient(InvocationContext.AuthenticationCredentialsProviders);
        var endpoint = "/drive/list/items" +
                        "?$select=id" +
                        "&$expand=driveItem($select=id,name,parentReference)" +
                        "&$filter=startswith(fields/ContentTypeId,'0x0120')" + 
                        "&$top=1000";
        var foldersDictionary = new Dictionary<string, string>();
        var foldersAmount = 0;

        do
        {
            var request = new SharePointRequest(endpoint, Method.Get, InvocationContext.AuthenticationCredentialsProviders);
            request.AddHeader("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly");
            var folders = await client.ExecuteWithHandling<ListWrapper<DriveItemWrapper<FolderMetadataDto>>>(request);
            var filteredFolders = folders.Value
                .Select(w => w.DriveItem)
                .Select(i => new { i.Id, Path = GetFolderPath(i) })
                .Where(i => string.IsNullOrEmpty(context.SearchString) ||
                           i.Path.Contains(context.SearchString, StringComparison.OrdinalIgnoreCase));

            foreach (var file in filteredFolders)
                foldersDictionary.Add(file.Id, file.Path);

            foldersAmount += filteredFolders.Count();
            endpoint = folders.ODataNextLink;
            if (endpoint != null && !endpoint.StartsWith("/drive/list/items"))
            {
                var queryString = endpoint.Contains("?") ? endpoint.Substring(endpoint.IndexOf("?")) : "";
                endpoint = "/drive/list/items" + queryString;
            }
            if (endpoint != null && !Uri.IsWellFormedUriString(endpoint, UriKind.Relative))
            {
                endpoint = null;
            }
        } while (foldersAmount < 20 && endpoint != null);

        foreach (var folder in foldersDictionary)
        {
            var folderPath = folder.Value;
            if (folderPath.Length > 40)
            {
                var folderPathParts = folderPath.Split("/");
                if (folderPathParts.Length > 3)
                {
                    folderPath = string.Join("/", folderPathParts[1], "...", folderPathParts[^2], folderPathParts[^1]);
                    foldersDictionary[folder.Key] = folderPath;
                }
            }
        }

        string rootName = "Root Folder (default)";
        var request1 = new SharePointRequest("/drive/root", Method.Get, InvocationContext.AuthenticationCredentialsProviders);
        var rootFolder = await client.ExecuteWithHandling<FolderMetadataDto>(request1);
        rootName = rootFolder.Name ?? "Root Folder (unnamed)";
        if (string.IsNullOrWhiteSpace(context.SearchString) ||
            rootName.Contains(context.SearchString, StringComparison.OrdinalIgnoreCase))
        {
            foldersDictionary.Add(rootFolder.Id, rootName);
        }

        return foldersDictionary;
    }

    private string GetFolderPath(FolderMetadataDto folder)
    {
        if (folder.ParentReference == null || string.IsNullOrEmpty(folder.ParentReference.Path))
            return folder.Name;

        var parentPath = folder.ParentReference.Path.Split("root:");
        return parentPath.Length > 1 && !string.IsNullOrEmpty(parentPath[1])
            ? $"{parentPath[1].TrimStart('/')}/{folder.Name}"
            : folder.Name;
    }
}