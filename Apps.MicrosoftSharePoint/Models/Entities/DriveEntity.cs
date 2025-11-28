using Newtonsoft.Json;

namespace Apps.MicrosoftSharePoint.Models.Entities;

public class DriveEntity
{
    [JsonProperty("id")]
    public string Id { get; set; }

    [JsonProperty("name")]
    public string Name { get; set; }

    [JsonProperty("lastModifiedDateTime")]
    public DateTime? LastModified { get; set; }
}
