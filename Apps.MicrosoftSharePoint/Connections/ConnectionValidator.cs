using Apps.MicrosoftSharePoint.Api;
using Blackbird.Applications.Sdk.Common.Authentication;
using Blackbird.Applications.Sdk.Common.Connections;
using RestSharp;

namespace Apps.MicrosoftSharePoint.Connections;

public class ConnectionValidator : IConnectionValidator
{
    public async ValueTask<ConnectionValidationResponse> ValidateConnection(
        IEnumerable<AuthenticationCredentialsProvider> authenticationCredentialsProviders, 
        CancellationToken cancellationToken)
    {
        var siteId = authenticationCredentialsProviders.First(p => p.KeyName == "SiteId").Value;
        
        if (siteId == null) 
            return new ConnectionValidationResponse
            {
                IsValid = false,
                Message = "Please provide correct site name."
            };
            
        var client = new MicrosoftSharePointClient(authenticationCredentialsProviders);
        var request = new MicrosoftSharePointRequest("", Method.Get);
        
        try
        {
            await client.ExecuteWithErrorHandling(request);
            return new ConnectionValidationResponse
            {
                IsValid = true,
                Message = "Success"
            };
        }
        catch (Exception)
        {
            return new ConnectionValidationResponse
            {
                IsValid = false,
                Message = "Ping failed"
            };
        }
    }
}