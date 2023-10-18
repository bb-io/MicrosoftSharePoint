using Apps.MicrosoftSharePoint.DataSourceHandlers;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Dynamic;
using File = Blackbird.Applications.Sdk.Common.Files.File;

namespace Apps.MicrosoftSharePoint.Models.Requests;

public class UploadFileRequest
{
    public File File { get; set; }
    
    [Display("Conflict behavior")]
    [DataSource(typeof(ConflictBehaviorDataSourceHandler))]
    public string ConflictBehavior { get; set; }
}