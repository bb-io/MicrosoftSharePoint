using System.Net;
using Polly;
using Polly.Retry;
using RestSharp;

namespace Apps.MicrosoftSharePoint.Api;

public static class SharePointRetryPolicy
{
    private const int BaseBackoffSeconds = 1;
    private const int MaxBackoffSeconds = 16;
    private const int MaxRetryAfterSeconds = 60;

    private static readonly HttpStatusCode[] TransientStatusCodes =
    {
        HttpStatusCode.TooManyRequests,
        HttpStatusCode.InternalServerError,
        HttpStatusCode.BadGateway,
        HttpStatusCode.ServiceUnavailable,
        HttpStatusCode.GatewayTimeout
    };

    public static AsyncRetryPolicy<RestResponse> Create(int retryCount) => Policy
        .HandleResult<RestResponse>(response => TransientStatusCodes.Contains(response.StatusCode))
        .WaitAndRetryAsync(retryCount,
            (retryAttempt, result, _) => GetRetryDelay(retryAttempt, result.Result),
            (_, _, _, _) => Task.CompletedTask);

    private static TimeSpan GetRetryDelay(int retryAttempt, RestResponse? response)
    {
        if (TryGetRetryAfterSeconds(response, out var retryAfterSeconds))
            return TimeSpan.FromSeconds(Math.Min(retryAfterSeconds, MaxRetryAfterSeconds));

        var backoff = Math.Min(BaseBackoffSeconds * Math.Pow(2, retryAttempt - 1), MaxBackoffSeconds);
        return TimeSpan.FromSeconds(backoff) + TimeSpan.FromMilliseconds(Random.Shared.Next(0, 500));
    }

    private static bool TryGetRetryAfterSeconds(RestResponse? response, out int seconds)
    {
        seconds = 0;

        var retryAfter = response?.Headers?
            .FirstOrDefault(header => string.Equals(header.Name, "Retry-After", StringComparison.OrdinalIgnoreCase))?
            .Value?.ToString();

        return int.TryParse(retryAfter, out seconds) && seconds > 0;
    }
}
