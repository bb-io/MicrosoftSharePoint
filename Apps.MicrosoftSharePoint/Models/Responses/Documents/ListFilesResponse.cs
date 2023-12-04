using Apps.MicrosoftSharePoint.Models.Dtos.Documents;

namespace Apps.MicrosoftSharePoint.Models.Responses.Documents;

public class ListFilesResponse
{
    public IEnumerable<FileMetadataDto> Files { get; set; }
}