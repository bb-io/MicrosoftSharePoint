using Apps.MicrosoftSharePoint.Dtos;

namespace Apps.MicrosoftSharePoint.Models.Responses;

public record ListFilesResponse(IEnumerable<FileMetadataDto> Files);