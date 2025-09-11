using Newtonsoft.Json;
using Blackbird.Applications.Sdk.Common;

namespace Apps.MicrosoftSharePoint.Dtos;

public class FieldDto
{
    public string? Title { get; set; }

    [JsonIgnore]
    private string? ImageTags { get; set; }

    [JsonProperty("ImageTags")]
    [Display("Image tags")]
    public string[] ImageTagArray => string.IsNullOrEmpty(ImageTags)
        ? Array.Empty<string>()
        : ImageTags.Split(';', StringSplitOptions.RemoveEmptyEntries);
}
