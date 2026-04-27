using System.Text.RegularExpressions;

namespace Apps.MicrosoftSharePoint.Extensions;
public static class StringExtensions
{
    public static string SanitizeFileName(this string input)
    {
        if (string.IsNullOrWhiteSpace(input))
            return string.Empty;

        // Replace invalid SharePoint characters with underscore
        string invalidCharsPattern = @"[""*:<>\?\/\\|#%]";
        string sanitized = Regex.Replace(input, invalidCharsPattern, "_");

        sanitized = sanitized.Trim();
        sanitized = sanitized.TrimEnd('.');

        return sanitized;
    }
}
