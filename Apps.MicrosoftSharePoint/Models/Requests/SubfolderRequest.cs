using Blackbird.Applications.Sdk.Common;

namespace Apps.MicrosoftSharePoint.Models.Requests;

public class SubfolderRequest
{
    [Display("Include subfolders", Description = "Default is false")]
    public bool? IncludeSubfolders { get; set; }

    public SubfolderRequest ApplyDefaultValues()
    {
        IncludeSubfolders ??= false;
        return this;
    }
}
