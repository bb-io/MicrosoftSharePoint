using Apps.MicrosoftSharePoint.Api;
using Apps.MicrosoftSharePoint.Models.Dtos;
using Apps.MicrosoftSharePoint.Models.Dtos.SitePages;
using Blackbird.Applications.Sdk.Common.Dynamic;
using Blackbird.Applications.Sdk.Common.Invocation;
using RestSharp;

namespace Apps.MicrosoftSharePoint.DataSourceHandlers;

public class NewsPostDataSourceHandler : MicrosoftSharePointInvocable, IAsyncDataSourceHandler
{
    public NewsPostDataSourceHandler(InvocationContext invocationContext) : base(invocationContext, isBetaApi: true)
    {
    }

    public async Task<Dictionary<string, string>> GetDataAsync(DataSourceContext context,
        CancellationToken cancellationToken)
    {
        const int limit = 20;
        var endpoint = $"/pages/microsoft.graph.sitePage?filter=promotionKind eq 'newsPost'&select=id,name,title,description,promotionKind&top={limit}";
        var postsDictionary = new Dictionary<string, string>();
        var postsAmount = 0;
        
        do
        {
            var request = new MicrosoftSharePointRequest(endpoint, Method.Get);
            var response = await Client.ExecuteWithErrorHandling<ListWrapper<BaseSitePageDto>>(request);
            var filteredPosts = response.Value
                .Where(post => context.SearchString == null 
                               || post.Name.Contains(context.SearchString, StringComparison.OrdinalIgnoreCase)
                               || (post.Title != null 
                                   && post.Title.Contains(context.SearchString, StringComparison.OrdinalIgnoreCase))
                               || (post.Description != null 
                                   && post.Description.Contains(context.SearchString, StringComparison.OrdinalIgnoreCase)))
                .ToArray();
            
            foreach (var post in filteredPosts)
                postsDictionary.Add(post.Id, post.Title ?? post.Name);

            postsAmount += filteredPosts.Length;
            endpoint = response.ODataNextLink == null ? null : "/pages" + response.ODataNextLink?.Split("pages")[^1];
        } while (postsAmount < limit && endpoint != null);

        return postsDictionary;
    }
}