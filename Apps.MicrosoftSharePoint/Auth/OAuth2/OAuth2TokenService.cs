﻿using Blackbird.Applications.Sdk.Common.Authentication.OAuth2;
using Newtonsoft.Json;

namespace Apps.MicrosoftSharePoint.Auth.OAuth2;

public class OAuth2TokenService : IOAuth2TokenService
{ 
    private const string TokenUrl = "https://login.microsoftonline.com/common/oauth2/v2.0/token";
    private const string ExpiresAtKeyName = "expires_at";
    
    public bool IsRefreshToken(Dictionary<string, string> values) 
        => values.TryGetValue(ExpiresAtKeyName, out var expireValue) && DateTime.UtcNow > DateTime.Parse(expireValue);
    
    public async Task<Dictionary<string, string>> RefreshToken(Dictionary<string, string> values, 
        CancellationToken cancellationToken) 
    { 
        const string grantType = "refresh_token";
        var bodyParameters = new Dictionary<string, string>
        {
            { "grant_type", grantType },
            { "refresh_token", values["refresh_token"] },
            { "client_id", ApplicationConstants.ClientId },
            { "client_secret", ApplicationConstants.ClientSecret }
        };
        return await RequestToken(bodyParameters, cancellationToken);
    }
    
    public async Task<Dictionary<string, string?>> RequestToken(string state, string code, 
        Dictionary<string, string> values, CancellationToken cancellationToken)
    { 
        const string grantType = "authorization_code"; 
        var bodyParameters = new Dictionary<string, string> 
        { 
            { "code", code },
            { "grant_type", grantType },
            { "client_id", ApplicationConstants.ClientId }, 
            { "client_secret", ApplicationConstants.ClientSecret },
            { "redirect_uri", ApplicationConstants.RedirectUri }
        };
        return await RequestToken(bodyParameters, cancellationToken);
    }

    public Task RevokeToken(Dictionary<string, string> values)
    { 
        throw new NotImplementedException();
    }
    
    private async Task<Dictionary<string, string>> RequestToken(Dictionary<string, string> bodyParameters, 
        CancellationToken cancellationToken)
    { 
        var utcNow = DateTime.UtcNow;
        using HttpClient httpClient = new HttpClient(); 
        httpClient.DefaultRequestHeaders.Add("Accept", "application/json"); 
        using var httpContent = new FormUrlEncodedContent(bodyParameters); 
        using var response = await httpClient.PostAsync(TokenUrl, httpContent, cancellationToken); 
        var responseContent = await response.Content.ReadAsStringAsync(cancellationToken); 
        var resultDictionary = JsonConvert.DeserializeObject<Dictionary<string, object>>(responseContent)?
                                   .ToDictionary(r => r.Key, r => r.Value?.ToString()) 
                               ?? throw new InvalidOperationException($"Invalid response content: {responseContent}");
        var expiresIn = int.Parse(resultDictionary["expires_in"]);
        var expiresAt = utcNow.AddSeconds(expiresIn);
        resultDictionary.Add(ExpiresAtKeyName, expiresAt.ToString());
        return resultDictionary;
    }
}