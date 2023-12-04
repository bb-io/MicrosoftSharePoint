using Apps.MicrosoftSharePoint.Models.Dtos.Documents;

namespace Apps.MicrosoftSharePoint.Models.Responses.Documents;

public class ListFoldersResponse
{
    public IEnumerable<FolderMetadataDto> Folders { get; set; }
}