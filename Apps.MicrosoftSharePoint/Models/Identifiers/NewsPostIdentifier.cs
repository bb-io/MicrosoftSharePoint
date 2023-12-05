using Apps.MicrosoftSharePoint.DataSourceHandlers;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Dynamic;

namespace Apps.MicrosoftSharePoint.Models.Identifiers;

public class NewsPostIdentifier
{
    [Display("Site page")]
    [DataSource(typeof(NewsPostDataSourceHandler))]
    public string SitePageId { get; set; }
}