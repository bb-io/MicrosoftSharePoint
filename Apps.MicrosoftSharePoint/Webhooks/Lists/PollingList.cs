using Apps.MicrosoftSharePoint.Models.Entities;
using Apps.MicrosoftSharePoint.Models.Responses;
using Apps.MicrosoftSharePoint.Webhooks.Memory;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Invocation;
using Blackbird.Applications.Sdk.Common.Polling;
using RestSharp;

namespace Apps.MicrosoftSharePoint.Webhooks.Lists
{
    [PollingEventList]
    public class PollingList : BaseInvocable
    {
        public PollingList(InvocationContext invocationContext) : base(invocationContext)
        {
        }

        [PollingEvent("On pages created or updated", "On pages created or updated")]
        public async Task<PollingEventResponse<PagesMemory, List<PageEntity>>> OnPagesCreatedOrUpdated(
            PollingEventRequest<PagesMemory> request)
        {
            var allPages = await ListAllPages();
            var newPagesState = allPages.Value.Select(x => $"{x.Id}|{x.LastModifiedDateTime}").ToList();
            if (request.Memory == null)
            {
                return new()
                {
                    FlyBird = false,
                    Memory = new PagesMemory() { PagesState = newPagesState }
                };
            }
            var changedItems = newPagesState.Except(request.Memory.PagesState).ToList();
            if (changedItems.Count == 0)
                return new()
                {
                    FlyBird = false,
                    Memory = new PagesMemory() { PagesState = newPagesState }
                };
            var changedPagesId = changedItems.Select(x => x.Split('|').First()).ToList();
            return new()
            {
                FlyBird = true,
                Memory = new PagesMemory() { PagesState = newPagesState },
                Result = allPages.Value.Where(x => changedPagesId.Contains(x.Id)).ToList()
            };
        }

        [PollingEvent("On pages deleted", "On pages deleted")]
        public async Task<PollingEventResponse<PagesMemory, List<PageEntity>>> OnPagesDeleted(
            PollingEventRequest<PagesMemory> request)
        {
            var allPages = await ListAllPages();
            var newPagesState = allPages.Value.Select(x => $"{x.Id}|{x.LastModifiedDateTime}").ToList();
            if (request.Memory == null)
            {
                return new()
                {
                    FlyBird = false,
                    Memory = new PagesMemory() { PagesState = newPagesState }
                };
            }
            var deletedPagesIds = request.Memory.PagesState.Except(newPagesState).ToList();
            if (deletedPagesIds.Count == 0)
                return new()
                {
                    FlyBird = false,
                    Memory = new PagesMemory() { PagesState = newPagesState }
                };
            return new()
            {
                FlyBird = true,
                Memory = new PagesMemory() { PagesState = newPagesState },
                Result = allPages.Value.Where(x => deletedPagesIds.Contains(x.Id)).ToList()
            };
        }

        private async Task<ListResponse<PageEntity>> ListAllPages()
        {
            var client = new MicrosoftSharePointClient(InvocationContext.AuthenticationCredentialsProviders);
            var request =
                new MicrosoftSharePointRequest("pages", Method.Get, InvocationContext.AuthenticationCredentialsProviders);
            var response = await client.ExecuteWithHandling<ListResponse<PageEntity>>(request);
            return response;
        }
    }
}
