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
            
        var client = new SharePointBetaClient(authenticationCredentialsProviders);
        var request = new SharePointRequest("", Method.Get, authenticationCredentialsProviders);
        
        try
        {
            await client.ExecuteWithHandling(request);
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