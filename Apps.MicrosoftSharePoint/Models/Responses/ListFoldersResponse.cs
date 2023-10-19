using Apps.MicrosoftSharePoint.Dtos;

namespace Apps.MicrosoftSharePoint.Models.Responses;

public class ListFoldersResponse
{
    public IEnumerable<FolderMetadataDto> Folders { get; set; }
}