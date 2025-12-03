namespace Apps.MicrosoftSharePoint.Dtos;

public class ItemLocationDto
{
    public string? DriveId { get; set; }
    public string ItemId { get; set; } = string.Empty;
    public bool IsDefaultDrive { get; }

    public ItemLocationDto(string itemId)
    {
        DriveId = null;
        ItemId = itemId;
        IsDefaultDrive = true;
    }

    public ItemLocationDto(string driveId, string itemId)
    {
        DriveId = driveId;
        ItemId = itemId;
        IsDefaultDrive = false;
    }
}
