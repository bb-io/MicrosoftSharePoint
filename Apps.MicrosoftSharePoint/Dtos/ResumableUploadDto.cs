﻿namespace Apps.MicrosoftSharePoint.Dtos;

public class ResumableUploadDto
{
    public IEnumerable<string>? NextExpectedRanges { get; set; }
    public string? UploadUrl { get; set; }
}