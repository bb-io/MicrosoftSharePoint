using Blackbird.Applications.Sdk.Common;
using Apps.MicrosoftSharePoint.DataSourceHandlers;
using Blackbird.Applications.SDK.Blueprints.Interfaces.FileStorage;
using Blackbird.Applications.SDK.Extensions.FileManagement.Models.FileDataSourceItems;

namespace Apps.MicrosoftSharePoint.Models.Identifiers;

public class FileIdentifier : IDownloadFileInput
{
    [Display("File ID")] 
    [FileDataSource(typeof(FilePickerDataSourceHandler))]
    public string FileId { get; set; }

    [Display("Drive ID")]
    [FileDataSource(typeof(DriveDataSourceHandler))]
    public string? DriveId { get; set; }
}