using RestSharp;

namespace Apps.MicrosoftSharePoint.Api;

public class MicrosoftSharePointRequest : RestRequest
{
    public MicrosoftSharePointRequest(string endpoint, Method method) : base(endpoint, method)
    {
    }
}