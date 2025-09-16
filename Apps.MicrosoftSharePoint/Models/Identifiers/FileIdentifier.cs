using Apps.MicrosoftSharePoint.DataSourceHandlers;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Dynamic;
using Blackbird.Applications.SDK.Blueprints.Interfaces.FileStorage;

namespace Apps.MicrosoftSharePoint.Models.Identifiers;

public class FileIdentifier : IDownloadFileInput
{
    [Display("File ID")] 
    [DataSource(typeof(FileDataSourceHandler))]
    public string FileId { get; set; }
}