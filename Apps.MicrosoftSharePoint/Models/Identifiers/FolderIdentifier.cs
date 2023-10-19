using Apps.MicrosoftSharePoint.DataSourceHandlers;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Dynamic;

namespace Apps.MicrosoftSharePoint.Models.Identifiers;

public class FolderIdentifier
{
    [Display("Folder")]
    [DataSource(typeof(FolderDataSourceHandler))]
    public string FolderId { get; set; }
}

public class ParentFolderIdentifier
{
    [Display("Parent folder")]
    [DataSource(typeof(FolderDataSourceHandler))]
    public string ParentFolderId { get; set; }
}