using Newtonsoft.Json;
using Blackbird.Applications.Sdk.Common;
using Apps.MicrosoftSharePoint.Converters;

namespace Apps.MicrosoftSharePoint.Dtos;

public class FieldDto
{
    public string? Title { get; set; }

    [JsonProperty("MediaServiceImageTags")]
    [JsonConverter(typeof(ImageTagConverter))]
    [Display("Image tags")]
    public string[]? ImageTags { get; set; }
}
