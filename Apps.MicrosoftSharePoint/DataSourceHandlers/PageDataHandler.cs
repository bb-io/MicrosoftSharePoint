using System.Web;
using Apps.MicrosoftSharePoint.Models.Entities;
using Apps.MicrosoftSharePoint.Models.Requests.Pages;
using Apps.MicrosoftSharePoint.Models.Responses;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Dynamic;
using Blackbird.Applications.Sdk.Common.Invocation;
using RestSharp;

namespace Apps.MicrosoftSharePoint.DataSourceHandlers;

public class PageDataHandler : BaseInvocable, IAsyncDataSourceHandler
{
    private readonly PageRequest _pageRequest;

    public PageDataHandler(InvocationContext invocationContext, [ActionParameter] PageRequest pageRequest) : base(
        invocationContext)
    {
        _pageRequest = pageRequest;
    }

    public async Task<Dictionary<string, string>> GetDataAsync(DataSourceContext context,
        CancellationToken cancellationToken)
    {
        var client = new SharePointBetaClient(InvocationContext.AuthenticationCredentialsProviders);

        var request =
            new SharePointRequest("pages", Method.Get, InvocationContext.AuthenticationCredentialsProviders);
        var response = await client.ExecuteWithHandling<ListResponse<PageEntity>>(request);

        var pages = response.Value
            .Where(x => context.SearchString is null ||
                        x.Name.Contains(context.SearchString, StringComparison.OrdinalIgnoreCase))
            .OrderByDescending(x => x.LastModifiedDateTime);

        var pagesCount = 50;
        if (!string.IsNullOrWhiteSpace(_pageRequest.Locale))
            return pages
                .Where(x => x.WebUrl.Split("/").SkipLast(1).Last() == _pageRequest.Locale)
                .Take(pagesCount)
                .ToDictionary(x => x.Id, x => x.Name);

        return pages
            .Take(pagesCount)
            .ToDictionary(x => x.Id, x => GetPagePath(x.WebUrl));
    }

    private string GetPagePath(string url)
    {
        var segments = url.Split("/");
        var name = string.Join("/", segments[^2].Trim('/'), segments[^1]);

        var result = name.StartsWith("SitePages/") ? name.Substring("SitePages/".Length) : name;
        return HttpUtility.UrlDecode(result);
    }
}