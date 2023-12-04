using Blackbird.Applications.Sdk.Common;

namespace Apps.MicrosoftSharePoint.Models.Dtos;

public class UserDto
{
    [Display("User ID")]
    public string Id { get; set; }
    
    public string Email { get; set; }
    
    [Display("Display name")]
    public string DisplayName { get; set; }
}