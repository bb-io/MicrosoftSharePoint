using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Files;
using Blackbird.Applications.Sdk.Common.Dictionaries;
using Blackbird.Applications.SDK.Blueprints.Interfaces.FileStorage;
using Apps.MicrosoftSharePoint.DataSourceHandlers;

namespace Apps.MicrosoftSharePoint.Models.Requests;

public class UploadFileRequest : IUploadFileInput
{
    public FileReference File { get; set; }
    
    [Display("Conflict behavior")]
    [StaticDataSource(typeof(ConflictBehaviorDataSourceHandler))]
    public string ConflictBehavior { get; set; }
}