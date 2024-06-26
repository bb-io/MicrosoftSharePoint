using Apps.MicrosoftSharePoint.Models.Entities;

namespace Apps.MicrosoftSharePoint.Webhooks.Payload
{
    public class PollingPagesUpdatedResponse
    {
        public List<PageEntity> Pages { get; set; }
    }
}
