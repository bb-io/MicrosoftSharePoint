using Apps.MicrosoftSharePoint.Dtos;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Dynamic;
using Blackbird.Applications.Sdk.Common.Invocation;
using RestSharp;

namespace Apps.MicrosoftSharePoint.DataSourceHandlers;

public class FileDataSourceHandler : BaseInvocable, IAsyncDataSourceItemHandler
{
    public FileDataSourceHandler(InvocationContext invocationContext) : base(invocationContext)
    {
    }

    async Task<IEnumerable<DataSourceItem>> IAsyncDataSourceItemHandler.GetDataAsync(DataSourceContext context, CancellationToken cancellationToken)
    {
        var client = new SharePointBetaClient(InvocationContext.AuthenticationCredentialsProviders);
        var endpoint = "/drive/list/items?$select=id&$expand=driveItem($select=id,name,parentReference)&" +
                       "$filter=fields/ContentType eq 'Document'&$top=1000";
        var filesDictionary = new List<DataSourceItem>();
        var filesAmount = 0;

        do
        {
            var request = new SharePointRequest(endpoint, Method.Get,
                InvocationContext.AuthenticationCredentialsProviders);
            request.AddHeader("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly");
            var files = await client.ExecuteWithHandling<ListWrapper<DriveItemWrapper<FileMetadataDto>>>(request);
            var filteredFiles = files.Value
                .Select(w => w.DriveItem)
                .Select(i => new { i.Id, Path = GetFilePath(i) })
                .Where(i => i.Path.Contains(context.SearchString, StringComparison.OrdinalIgnoreCase));

            foreach (var file in filteredFiles)
                filesDictionary.Add(new DataSourceItem(file.Id, file.Path));

            filesAmount += filteredFiles.Count();
            endpoint = files.ODataNextLink == null ? null : "/drive" + files.ODataNextLink?.Split("drive")[^1];
        } while (filesAmount < 20 && endpoint != null);

        foreach (var file in filesDictionary)
        {
            var filePath = file.Value;
            if (filePath.Length > 40)
            {
                var filePathParts = filePath.Split("/");
                if (filePathParts.Length > 3)
                {
                    filePath = string.Join("/", filePathParts[0], "...", filePathParts[^2], filePathParts[^1]);
                    filesDictionary.First( x => x.Value == file.Value).DisplayName = filePath;
                }
            }
        }
        return filesDictionary;
    }

    private string GetFilePath(FileMetadataDto file)
    {
        var parentPath = file.ParentReference.Path.Split("root:");
        if (parentPath[1] == "")
            return file.Name;

        return $"{parentPath[1].Substring(1)}/{file.Name}";
    }
}