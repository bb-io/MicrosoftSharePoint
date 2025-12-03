using Apps.MicrosoftSharePoint.Dtos;

namespace Apps.MicrosoftSharePoint.Helper;

public static class ItemIdParser
{
    public static ItemLocationDto Parse(string? inputId)
    {
        if (string.IsNullOrEmpty(inputId))
            return new ItemLocationDto("root");

        if (inputId.Contains('#'))
        {
            var parts = inputId.Split('#');
            var itemId = parts.Length > 1 ? parts[1] : "root";
            return new ItemLocationDto(parts[0], itemId);
        }

        return new ItemLocationDto(inputId);
    }

    public static string Format(string driveId, string itemId, string defaultDriveId)
    {
        if (string.Equals(driveId, defaultDriveId, StringComparison.OrdinalIgnoreCase))
            return itemId;

        return $"{driveId}#{itemId}";
    }
}
