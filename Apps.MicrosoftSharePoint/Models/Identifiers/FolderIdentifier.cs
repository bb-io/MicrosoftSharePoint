using Apps.MicrosoftSharePoint.DataSourceHandlers;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Dynamic;

namespace Apps.MicrosoftSharePoint.Models.Identifiers;

public class FolderIdentifier
{
    [Display("Folder ID", Description = "Enter the unique ID of the folder. For example: '01C7WXPSBPDJQQ2E2CTBFI5XXXXXXXXXX'.")]
    [DataSource(typeof(FolderDataSourceHandler))]
    public string FolderId { get; set; }
}

public class ParentFolderIdentifier
{
    [Display("Parent folder ID", Description = "Enter the unique ID of the folder. For example: '01C7WXPSBPDJQQ2E2CTBFI5XXXXXXXXXX'.")]
    [DataSource(typeof(FolderDataSourceHandler))]
    public string ParentFolderId { get; set; }
}