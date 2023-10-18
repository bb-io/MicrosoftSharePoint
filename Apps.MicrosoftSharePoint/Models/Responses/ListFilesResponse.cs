using Apps.MicrosoftSharePoint.Dtos;

namespace Apps.MicrosoftSharePoint.Models.Responses;

public class ListFilesResponse
{
    public IEnumerable<FileMetadataDto> Files { get; set; }
}