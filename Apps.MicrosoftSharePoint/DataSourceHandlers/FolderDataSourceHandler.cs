using Apps.MicrosoftSharePoint.Api;
using Apps.MicrosoftSharePoint.Models.Dtos;
using Apps.MicrosoftSharePoint.Models.Dtos.Documents;
using Blackbird.Applications.Sdk.Common.Dynamic;
using Blackbird.Applications.Sdk.Common.Invocation;
using RestSharp;

namespace Apps.MicrosoftSharePoint.DataSourceHandlers;

public class FolderDataSourceHandler : MicrosoftSharePointInvocable, IAsyncDataSourceHandler
{
    public FolderDataSourceHandler(InvocationContext invocationContext) : base(invocationContext)
    {
    }

    public async Task<Dictionary<string, string>> GetDataAsync(DataSourceContext context,
        CancellationToken cancellationToken)
    {
        var endpoint = "/drive/list/items?$select=id&$expand=driveItem($select=id,name,parentReference)&" +
                       "$filter=fields/ContentType eq 'Folder'&$top=20";
        var foldersDictionary = new Dictionary<string, string>();
        var foldersAmount = 0;

        do
        {
            var request = new MicrosoftSharePointRequest(endpoint, Method.Get);
            request.AddHeader("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly");
            var folders = await Client.ExecuteWithErrorHandling<ListWrapper<DriveItemWrapper<FolderMetadataDto>>>(request);
            var filteredFolders = folders.Value
                .Select(w => w.DriveItem)
                .Select(i => new { i.Id, Path = GetFolderPath(i) })
                .Where(i => i.Path.Contains(context.SearchString, StringComparison.OrdinalIgnoreCase));
            
            foreach (var file in filteredFolders)
                foldersDictionary.Add(file.Id, file.Path);
            
            foldersAmount += filteredFolders.Count();
            endpoint = folders.ODataNextLink == null ? null : "/drive" + folders.ODataNextLink?.Split("drive")[^1];
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

        const string rootName = "My files (root folder)";
        if (string.IsNullOrWhiteSpace(context.SearchString) 
            || rootName.Contains(context.SearchString, StringComparison.OrdinalIgnoreCase))
        {
            var request = new MicrosoftSharePointRequest("/drive/root", Method.Get);
            var rootFolder = await Client.ExecuteWithErrorHandling<FolderMetadataDto>(request);
            foldersDictionary.Add(rootFolder.Id, rootName);
        }
            
        return foldersDictionary;
    }

    private string GetFolderPath(FolderMetadataDto folder)
    {
        var parentPath = folder.ParentReference.Path.Split("root:");
        if (parentPath[1] == "")
            return folder.Name;

        return $"{parentPath[1].Substring(1)}/{folder.Name}";
    }
}