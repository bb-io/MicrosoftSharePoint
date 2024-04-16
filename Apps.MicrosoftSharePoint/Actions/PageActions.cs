using System.Net.Mime;
using Apps.MicrosoftSharePoint.HtmlConversion;
using Apps.MicrosoftSharePoint.Models.Requests;
using Apps.MicrosoftSharePoint.Models.Requests.Pages;
using Apps.MicrosoftSharePoint.Models.Responses;
using Apps.MicrosoftSharePoint.Models.Responses.Pages;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Actions;
using Blackbird.Applications.Sdk.Common.Authentication;
using Blackbird.Applications.Sdk.Common.Invocation;
using Blackbird.Applications.SDK.Extensions.FileManagement.Interfaces;
using Blackbird.Applications.Sdk.Utils.Extensions.Http;
using Newtonsoft.Json.Linq;
using RestSharp;

namespace Apps.MicrosoftSharePoint.Actions;

[ActionList]
public class PageActions : BaseInvocable
{
    private readonly IEnumerable<AuthenticationCredentialsProvider> _authenticationCredentialsProviders;
    private readonly MicrosoftSharePointClient _client;
    private readonly IFileManagementClient _fileManagementClient;

    public PageActions(InvocationContext invocationContext, IFileManagementClient fileManagementClient)
        : base(invocationContext)
    {
        _authenticationCredentialsProviders = invocationContext.AuthenticationCredentialsProviders;
        _client = new MicrosoftSharePointClient(_authenticationCredentialsProviders);
        _fileManagementClient = fileManagementClient;
    }

    [Action("Get page as HTML", Description = "Get content of a specific page in HTML format")]
    public async Task<FileResponse> GetPageContent([ActionParameter] PageRequest pageRequest)
    {
        var pageContent = await GetPageJObject(pageRequest.PageId);
        var html = SharePointHtmlConverter.ToHtml(pageContent);

        return new()
        {
            File = await _fileManagementClient.UploadAsync(new MemoryStream(html), MediaTypeNames.Text.Html,
                $"{Path.GetFileNameWithoutExtension(pageContent.Name)}.html")
        };
    }

    [Action("Update page from HTML", Description = "Update content of a specific page from HTML file")]
    public async Task UpdatePageContent([ActionParameter] PageRequest pageRequest, [ActionParameter] FileRequest file)
    {
        var fileStream = await _fileManagementClient.DownloadAsync(file.File);
        var (title, content) = SharePointHtmlConverter.ToJson(fileStream);

        content.Descendants()
            .Where(x => x is JProperty jProperty && jProperty.Name.EndsWith("@odata.context"))
            .Cast<JProperty>()
            .ToList()
            .ForEach(x => x.Remove());

        var request = new MicrosoftSharePointRequest($"pages/{pageRequest.PageId}/microsoft.graph.sitePage",
                Method.Patch, _authenticationCredentialsProviders)
            .WithJsonBody(new
            {
                title,
                titleArea = new
                {
                    title
                },
                canvasLayout = content
            });
        await _client.ExecuteWithHandling(request);
    }

    private Task<PageContentResponse> GetPageJObject(string pageId)
    {
        var request = new MicrosoftSharePointRequest($"pages/{pageId}/microsoft.graph.sitepage?expand=canvasLayout",
            Method.Get, _authenticationCredentialsProviders);
        return _client.ExecuteWithHandling<PageContentResponse>(request);
    }
}