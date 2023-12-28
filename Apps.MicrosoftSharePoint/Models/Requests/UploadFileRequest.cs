using Apps.MicrosoftSharePoint.DataSourceHandlers;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Dynamic;
using Blackbird.Applications.Sdk.Common.Files;

namespace Apps.MicrosoftSharePoint.Models.Requests;

public class UploadFileRequest
{
    public FileReference File { get; set; }
    
    [Display("Conflict behavior")]
    [DataSource(typeof(ConflictBehaviorDataSourceHandler))]
    public string ConflictBehavior { get; set; }
}