using Apps.MicrosoftSharePoint.Api;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Invocation;

namespace Apps.MicrosoftSharePoint;

public class MicrosoftSharePointInvocable : BaseInvocable
{
    protected readonly MicrosoftSharePointClient Client;

    protected MicrosoftSharePointInvocable(InvocationContext invocationContext, bool isBetaApi = false) 
        : base(invocationContext)
    {
        Client = new MicrosoftSharePointClient(invocationContext.AuthenticationCredentialsProviders, isBetaApi);
    }
}