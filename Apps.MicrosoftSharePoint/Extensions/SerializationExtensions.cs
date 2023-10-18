using Newtonsoft.Json;

namespace Apps.MicrosoftSharePoint.Extensions;

public static class SerializationExtensions
{
    public static T DeserializeObject<T>(this string content)
    {
        var deserializedObject = JsonConvert.DeserializeObject<T>(content, new JsonSerializerSettings
            {
                MissingMemberHandling = MissingMemberHandling.Ignore,
                DateTimeZoneHandling = DateTimeZoneHandling.Local
            }
        );
        return deserializedObject;
    }
}