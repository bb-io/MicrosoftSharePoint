using Apps.MicrosoftSharePoint.Models.Dtos.SitePages;
using Blackbird.Applications.Sdk.Common.Authentication;

namespace Apps.MicrosoftSharePoint.HtmlHelpers;

public static class HtmlToNewsPostConverter
{
    public static NewsPostContentDto ConvertToNewsPost(this string html, 
        IEnumerable<AuthenticationCredentialsProvider> authenticationCredentialsProviders)
    {
        
    } 
}