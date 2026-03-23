using Apps.MicrosoftSharePoint.Dtos;

namespace Apps.MicrosoftSharePoint.Models.Responses;

public record ListFoldersResponse(IEnumerable<FolderMetadataDto> Folders);