using Apps.MicrosoftSharePoint.DataSourceHandlers;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Dynamic;

namespace Apps.MicrosoftSharePoint.Webhooks.Inputs;

public class ContentTypeInput
{
    [Display("Content type")]
    [DataSource(typeof(ContentTypeDataSourceHandler))]
    public string? ContentType { get; set; }
}