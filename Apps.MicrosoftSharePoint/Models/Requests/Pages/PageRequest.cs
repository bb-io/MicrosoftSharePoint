using Apps.MicrosoftSharePoint.DataSourceHandlers;
using Apps.MicrosoftSharePoint.DataSourceHandlers.StaticDataSourceHandlers;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Dynamic;

namespace Apps.MicrosoftSharePoint.Models.Requests.Pages;

public class PageRequest
{
    [Display("Page ID")]
    [DataSource(typeof(PageDataHandler))]
    public string PageId { get; set; }
    
    [DataSource(typeof(LanguageDataHandler))]
    public string? Locale { get; set; }
}