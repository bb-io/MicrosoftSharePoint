using Blackbird.Applications.Sdk.Common.Files;
using Blackbird.Applications.SDK.Blueprints.Interfaces.FileStorage;

namespace Apps.MicrosoftSharePoint.Models.Responses;

public class FileResponse : IDownloadFileOutput
{
    public FileReference File { get; set; }
}