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
    private readonly PagesRequest _pagesRequest;
    
    public PageDataHandler(InvocationContext invocationContext, [ActionParameter] PagesRequest pagesRequest) : base(
        invocationContext)
    {
        _pagesRequest = pagesRequest;
    }

    public async Task<Dictionary<string, string>> GetDataAsync(DataSourceContext context,
        CancellationToken cancellationToken)
    {
        var client = new MicrosoftSharePointClient(InvocationContext.AuthenticationCredentialsProviders);

        var request =
            new MicrosoftSharePointRequest("pages", Method.Get, InvocationContext.AuthenticationCredentialsProviders);
        var response = await client.ExecuteWithHandling<ListResponse<PageEntity>>(request);

        var pages = response.Value
            .Where(x => context.SearchString is null ||
                        x.Name.Contains(context.SearchString, StringComparison.OrdinalIgnoreCase))
            .OrderByDescending(x => x.LastModifiedDateTime);

        var pagesCount = 50;
        if (!string.IsNullOrWhiteSpace(_pagesRequest.Locale))
            return pages
                .Where(x => x.WebUrl.Split("/").SkipLast(1).Last() == _pagesRequest.Locale)
                .Take(pagesCount)
                .ToDictionary(x => x.Id, x => x.Name);
        
        return pages
            .Where(x => x.WebUrl.Split("/").SkipLast(1).Last() == _pagesRequest.Locale)
            .Take(pagesCount)
            .ToDictionary(x => x.Id, x => GetPagePath(x.WebUrl));
    }

    private string GetPagePath(string url)
    {
        var segments = url.Split("/");
        return string.Join("/", segments[^2].Trim('/'), segments[^1]);
    }
}