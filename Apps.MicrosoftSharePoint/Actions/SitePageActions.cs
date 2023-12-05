using System.Net.Mime;
using System.Text;
using Apps.MicrosoftSharePoint.Api;
using Apps.MicrosoftSharePoint.HtmlHelpers;
using Apps.MicrosoftSharePoint.Models;
using Apps.MicrosoftSharePoint.Models.Dtos.SitePages;
using Apps.MicrosoftSharePoint.Models.Identifiers;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Actions;
using Blackbird.Applications.Sdk.Common.Invocation;
using RestSharp;

namespace Apps.MicrosoftSharePoint.Actions;

[ActionList]
public class SitePageActions : MicrosoftSharePointInvocable
{
    public SitePageActions(InvocationContext invocationContext) : base(invocationContext, isBetaApi: true)
    {
    }

    [Action("Get news post as HTML file", Description = "Retrieve a news post as HTML file.")]
    public async Task<FileWrapper> GetNewsPostAsHtml([ActionParameter] NewsPostIdentifier newsPostIdentifier)
    {
        var request = new MicrosoftSharePointRequest(
            $"/pages/{newsPostIdentifier.SitePageId}/microsoft.graph.sitePage?expand=canvasLayout", Method.Get);
        request.AddHeader("Accept", "application/json;odata.metadata=none");
        var newsPostContent = await Client.ExecuteWithErrorHandling<NewsPostContentDto>(request);
        var postHtml = newsPostContent.ConvertToHtml();
        var resultHtml = $"<html><body>{postHtml}</body></html>";

        return new FileWrapper
        {
            File = new(Encoding.UTF8.GetBytes(resultHtml))
            {
                Name = $"{(newsPostContent.Title ?? newsPostContent.Name).Replace(" ", "_")}.html",
                ContentType = MediaTypeNames.Text.Html
            }
        };
    }

    [Action("Create news post from HTML file", Description = "Create a news post from HTML file.")]
    public async Task CreateNewsPostFromHtml([ActionParameter] FileWrapper file)
    {
        
    }
}