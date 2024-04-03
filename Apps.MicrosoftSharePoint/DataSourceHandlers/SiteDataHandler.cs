// using Apps.MicrosoftSharePoint.Models.Entities;
// using Apps.MicrosoftSharePoint.Models.Responses;
// using Blackbird.Applications.Sdk.Common;
// using Blackbird.Applications.Sdk.Common.Dynamic;
// using Blackbird.Applications.Sdk.Common.Invocation;
//
// namespace Apps.MicrosoftSharePoint.DataSourceHandlers;
//
// public class SiteDataHandler : BaseInvocable, IAsyncDataSourceHandler
// {
//     public SiteDataHandler(InvocationContext invocationContext) : base(invocationContext)
//     {
//     }
//
//     public async Task<Dictionary<string, string>> GetDataAsync(DataSourceContext context,
//         CancellationToken cancellationToken)
//     {
//         var client = new MicrosoftSharePointClient(InvocationContext.AuthenticationCredentialsProviders);
//         var response = await client.ExecuteWithHandling<ListResponse<SiteEntity>>(new("getAllSites"));
//
//         return response.Value
//             .Where(x => context.SearchString is null ||
//                         x.Name.Contains(context.SearchString, StringComparison.OrdinalIgnoreCase))
//             .OrderByDescending(x => x.CreatedDateTime)
//             .ToDictionary(x => x.Id, x => x.Name);
//     }
// }