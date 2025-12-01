using Blackbird.Applications.Sdk.Common.Invocation;
using Blackbird.Applications.SDK.Extensions.FileManagement.Interfaces;
using Blackbird.Applications.SDK.Extensions.FileManagement.Models.FileDataSourceItems;

namespace Apps.MicrosoftSharePoint.DataSourceHandlers;

public class FilePickerDataSourceHandler(InvocationContext invocationContext)
    : BaseFileFolderPicker(invocationContext), IAsyncFileDataSourceItemHandler
{
    public async Task<IEnumerable<FileDataItem>> GetFolderContentAsync(FolderContentDataSourceContext context, CancellationToken ct)
    {
        return await GetFolderContent(context.FolderId, true, false);
    }

    public async Task<IEnumerable<FolderPathItem>> GetFolderPathAsync(FolderPathDataSourceContext context, CancellationToken ct)
    {
        return await GetFolderPath(context.FileDataItemId);
    }
}