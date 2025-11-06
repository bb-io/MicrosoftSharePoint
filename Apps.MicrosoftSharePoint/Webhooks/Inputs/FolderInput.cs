using Apps.MicrosoftSharePoint.DataSourceHandlers;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Dynamic;
using Blackbird.Applications.SDK.Extensions.FileManagement.Models.FileDataSourceItems;

namespace Apps.MicrosoftSharePoint.Webhooks.Inputs;

public class FolderInput
{
    [Display("Folder ID")]
    [FileDataSource(typeof(FolderPickerDataSourceHandler))]
    public string? ParentFolderId { get; set; }
}