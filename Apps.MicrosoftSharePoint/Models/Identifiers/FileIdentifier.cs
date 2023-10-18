using Apps.MicrosoftSharePoint.DataSourceHandlers;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Dynamic;

namespace Apps.MicrosoftSharePoint.Models.Identifiers;

public class FileIdentifier
{
    [Display("File")] 
    [DataSource(typeof(FileDataSourceHandler))]
    public string FileId { get; set; }
}