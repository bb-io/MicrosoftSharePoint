using RestSharp;
using System.Web;
using Apps.MicrosoftSharePoint.Models.Entities;
using Apps.MicrosoftSharePoint.Models.Responses;
using Apps.MicrosoftSharePoint.Models.Requests.Pages;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Dynamic;
using Blackbird.Applications.Sdk.Common.Invocation;

namespace Apps.MicrosoftSharePoint.DataSourceHandlers;

public class PageDataHandler(
    InvocationContext invocationContext,
    [ActionParameter] PageRequest pageRequest)
    : BaseInvocable(invocationContext), IAsyncDataSourceItemHandler
{
    public async Task<IEnumerable<DataSourceItem>> GetDataAsync(DataSourceContext context, CancellationToken ct)
    {
        var creds = InvocationContext.AuthenticationCredentialsProviders;

        var client = new SharePointBetaClient(creds);
        var request = new SharePointRequest("pages", Method.Get, creds);

        var response = await client.ExecuteWithHandling<ListResponse<PageEntity>>(request);
        var pages = response.Value
            .Where(x => context.SearchString is null ||
                        x.Name.Contains(context.SearchString, StringComparison.OrdinalIgnoreCase))
            .OrderByDescending(x => x.LastModifiedDateTime);

        var pagesCount = 50;
        if (!string.IsNullOrWhiteSpace(pageRequest.Locale))
            return pages
                .Where(x => x.WebUrl.Split("/").SkipLast(1).Last() == pageRequest.Locale)
                .Take(pagesCount)
                .Select(x => new DataSourceItem(x.Id, x.Name));

        return pages
            .Take(pagesCount)
            .Select(x => new DataSourceItem(x.Id, GetPagePath(x.WebUrl)));
    }

    private static string GetPagePath(string url)
    {
        var segments = url.Split("/");
        var name = string.Join("/", segments[^2].Trim('/'), segments[^1]);

        var result = name.StartsWith("SitePages/") ? name.Substring("SitePages/".Length) : name;
        return HttpUtility.UrlDecode(result);
    }
}