using Apps.MicrosoftSharePoint.Dtos;
using Newtonsoft.Json;

namespace Apps.MicrosoftSharePoint.Converters;

public class FieldsConverter : JsonConverter
{
    private class JsonListItem
    {
        public FieldDto Fields { get; set; }
    }

    public override bool CanWrite => false;
    public override bool CanConvert(Type objectType) => false;

    public override object ReadJson(JsonReader reader, Type objectType, object existingValue, JsonSerializer serializer)
    {
        if (reader.TokenType == JsonToken.StartObject)
        {
            var listItem = serializer.Deserialize<JsonListItem>(reader);
            return listItem.Fields;
        }

        return null;
    }

    public override void WriteJson(JsonWriter writer, object value, JsonSerializer serializer)
    {
        throw new NotImplementedException();
    }
}