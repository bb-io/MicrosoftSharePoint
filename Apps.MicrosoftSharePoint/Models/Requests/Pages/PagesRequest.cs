using Apps.MicrosoftSharePoint.DataSourceHandlers;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Dynamic;

namespace Apps.MicrosoftSharePoint.Models.Requests.Pages;

public class PagesRequest
{
    [Display("Page ID")]
    [DataSource(typeof(PageDataHandler))]
    public string PageId { get; set; }
    
    public string? Locale { get; set; }
}