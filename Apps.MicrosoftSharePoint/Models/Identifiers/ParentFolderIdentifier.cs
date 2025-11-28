using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.SDK.Extensions.FileManagement.Models.FileDataSourceItems;
using Apps.MicrosoftSharePoint.DataSourceHandlers;

namespace Apps.MicrosoftSharePoint.Models.Identifiers;

public class ParentFolderIdentifier
{
    [Display("Parent folder ID", Description = "Enter the unique ID of the folder. For example: '01C7WXPSBPDJQQ2E2CTBFI5XXXXXXXXXX'.")]
    [FileDataSource(typeof(FolderPickerDataSourceHandler))]
    public string ParentFolderId { get; set; }
}