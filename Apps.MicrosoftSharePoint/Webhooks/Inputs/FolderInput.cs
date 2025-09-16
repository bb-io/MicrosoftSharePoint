using Apps.MicrosoftSharePoint.DataSourceHandlers;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Dynamic;

namespace Apps.MicrosoftSharePoint.Webhooks.Inputs;

public class FolderInput
{
    [Display("Folder ID")] 
    [DataSource(typeof(FolderDataSourceHandler))]
    public string? ParentFolderId { get; set; }
}