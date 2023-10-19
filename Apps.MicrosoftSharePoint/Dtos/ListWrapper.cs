using Newtonsoft.Json;

namespace Apps.MicrosoftSharePoint.Dtos;

public class ListWrapper<T>
{
    [JsonProperty("@odata.nextLink")]
    public string? ODataNextLink { get; set; }
    
    [JsonProperty("@odata.deltaLink")]
    public string? ODataDeltaLink { get; set; }

    public IEnumerable<T> Value { get; set; }
}