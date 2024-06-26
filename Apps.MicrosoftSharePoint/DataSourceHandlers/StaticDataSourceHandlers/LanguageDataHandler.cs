using Blackbird.Applications.Sdk.Utils.Sdk.DataSourceHandlers;

namespace Apps.MicrosoftSharePoint.DataSourceHandlers.StaticDataSourceHandlers;

public class LanguageDataHandler : EnumDataHandler
{
    protected override Dictionary<string, string> EnumValues => new()
    {
        { "ar", "Arabic" },
        { "az-latn", "Azeri (Azerbaijani)" },
        { "eu", "Basque" },
        { "bs-latn", "Bosnian Latin" },
        { "bg", "Bulgarian" },
        { "ca", "Catalan" },
        { "zh-chs", "Chinese Simplified" },
        { "zh-cht", "Chinese Traditional" },
        { "hr", "Croatian" },
        { "cs", "Czech" },
        { "da", "Danish" },
        { "prs", "Dari" },
        { "nl", "Dutch" },
        { "en", "English" },
        { "et", "Estonian" },
        { "fi", "Finnish" },
        { "fr", "French" },
        { "gl", "Galician" },
        { "de", "German" },
        { "el", "Greek" },
        { "he", "Hebrew" },
        { "hi", "Hindi" },
        { "hu", "Hungarian" },
        { "id", "Indonesian" },
        { "ga", "Irish" },
        { "it", "Italian" },
        { "ja", "Japanese" },
        { "kk", "Kazakh" },
        { "ko", "Korean" },
        { "lv", "Latvian" },
        { "lt", "Lithuanian" },
        { "mk", "Macedonian" },
        { "ms", "Malay - Malaysia" },
        { "nb", "Norwegian (Bokmal)" },
        { "pl", "Polish" },
        { "pt-br", "Portuguese-Brazil" },
        { "pt-pt", "Portuguese-Portugal" },
        { "ro", "Romanian" },
        { "ru", "Russian" },
        { "sr-cyrl", "Serbian - Cyrillic RS" },
        { "sr-latn", "Serbian - Latin RS" },
        { "sk", "Slovak" },
        { "sl", "Slovenian" },
        { "es", "Spanish" },
        { "sv", "Swedish" },
        { "th", "Thai" },
        { "tr", "Turkish" },
        { "uk", "Ukrainian" },
        { "vi", "Vietnamese" },
        { "cy", "Welsh" }
    };
}