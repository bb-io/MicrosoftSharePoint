using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace Apps.MicrosoftSharePoint.Converters;

public class ImageTagConverter : JsonConverter<string[]>
{
    public override string[] ReadJson(JsonReader reader, Type objectType, string[] existingValue, bool hasExistingValue, JsonSerializer serializer)
    {
        var arr = JArray.Load(reader);
        var list = new List<string>();
        foreach (var item in arr)
        {
            var label = item["Label"]?.ToString();
            if (!string.IsNullOrEmpty(label))
                list.Add(label);
        }
        return list.ToArray();
    }

    public override void WriteJson(JsonWriter writer, string[] value, JsonSerializer serializer)
    {
        serializer.Serialize(writer, value);
    }
}
