using Blackbird.Applications.Sdk.Common;

namespace Apps.MicrosoftSharePoint.Models.Identifiers;

public class FolderIdentifier
{
    [Display("Parent folder")]
    //[DataSource(typeof(FolderDataSourceHandler))]
    public string FolderId { get; set; }
}