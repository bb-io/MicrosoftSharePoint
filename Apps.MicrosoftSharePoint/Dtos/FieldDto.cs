using Newtonsoft.Json;

namespace Apps.MicrosoftSharePoint.Dtos;

public class FieldDto
{
    public string? Title { get; set; }

    [JsonIgnore]
    public string? ImageTags { get; set; }

    [JsonProperty("ImageTags")]
    public string[] ImageTagArray => string.IsNullOrEmpty(ImageTags)
        ? Array.Empty<string>()
        : ImageTags.Split(';', StringSplitOptions.RemoveEmptyEntries);
}
