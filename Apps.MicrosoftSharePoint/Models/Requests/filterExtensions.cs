using Blackbird.Applications.Sdk.Common;

namespace Apps.MicrosoftSharePoint.Models.Requests;

public class filterExtensions
{
    [Display("Include these extensions only")]
    public IEnumerable<string>? Extensions { get; set; }
}
