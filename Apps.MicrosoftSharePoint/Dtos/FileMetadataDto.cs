﻿using Apps.MicrosoftSharePoint.Converters;
using Blackbird.Applications.Sdk.Common;
using Newtonsoft.Json;

namespace Apps.MicrosoftSharePoint.Dtos;

public class FileMetadataDto
{
    [Display("File ID")]
    public string Id { get; set; }
    
    [Display("Filename")]
    public string Name { get; set; }
    
    [Display("Web URL")]
    public string? WebUrl { get; set; }
    
    [Display("Size in bytes")]
    public long? Size { get; set; }
    
    [JsonProperty("file")]
    [JsonConverter(typeof(MimeTypeConverter))]
    [Display("Mime type")]
    public string? MimeType { get; set; }
    
    [JsonConverter(typeof(UserConverter))]
    [Display("Created by")]
    public UserDto? CreatedBy { get; set; }
    
    [Display("Created date and time")]
    public DateTime? CreatedDateTime { get; set; }
    
    [JsonConverter(typeof(UserConverter))]
    [Display("Last modified by")]
    public UserDto? LastModifiedBy { get; set; }
    
    [Display("Last modified date and time")]
    public DateTime? LastModifiedDateTime { get; set; }
    
    [Display("Parent reference")]
    public ParentReferenceDto? ParentReference { get; set; }
}