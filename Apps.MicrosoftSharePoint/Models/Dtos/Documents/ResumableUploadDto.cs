namespace Apps.MicrosoftSharePoint.Models.Dtos.Documents;

public class ResumableUploadDto
{
    public IEnumerable<string>? NextExpectedRanges { get; set; }
    public string? UploadUrl { get; set; }
}