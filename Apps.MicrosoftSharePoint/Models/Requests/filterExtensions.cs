using Blackbird.Applications.Sdk.Common;

namespace Apps.MicrosoftSharePoint.Models.Requests;

public class FilterExtensions
{
    [Display("Include these extensions only")]
    public IEnumerable<string>? Extensions { get; set; }
}
