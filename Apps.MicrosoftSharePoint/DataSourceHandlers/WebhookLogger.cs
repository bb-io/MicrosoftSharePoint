using RestSharp;
using Blackbird.Applications.Sdk.Utils.Extensions.Http;

namespace Apps.MicrosoftSharePoint.DataSourceHandlers;

public static class WebhookLogger
{
    public async static Task Log(string url, object payload)
    {
        var request = new RestRequest(new Uri(url), Method.Post).WithJsonBody(payload);
        var client = new RestClient();
        await client.ExecuteAsync(request);
    }
}
