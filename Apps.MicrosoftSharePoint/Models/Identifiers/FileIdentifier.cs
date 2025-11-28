using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.SDK.Blueprints.Interfaces.FileStorage;
using Blackbird.Applications.SDK.Extensions.FileManagement.Models.FileDataSourceItems;
using Apps.MicrosoftSharePoint.DataSourceHandlers;

namespace Apps.MicrosoftSharePoint.Models.Identifiers;

public class FileIdentifier : IDownloadFileInput
{
    [Display("File ID")] 
    [FileDataSource(typeof(FilePickerDataSourceHandler))]
    public string FileId { get; set; }
}