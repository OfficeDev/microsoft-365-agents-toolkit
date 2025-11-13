using Microsoft.Kiota.Abstractions.Authentication;

namespace TravelAgent.Bot.Plugins;

public class StaticTokenProvider(string token) : IAccessTokenProvider
{
    public AllowedHostsValidator AllowedHostsValidator => new(["graph.microsoft.com"]);

    public Task<string> GetAuthorizationTokenAsync(
        Uri uri,
#pragma warning disable CS8632 // The annotation for nullable reference types should only be used in code within a '#nullable' annotations context.
        Dictionary<string, object>? additionalAuthenticationContext = null,
#pragma warning restore CS8632 // The annotation for nullable reference types should only be used in code within a '#nullable' annotations context.
        CancellationToken cancellationToken = default)
    {
        return AllowedHostsValidator.AllowedHosts.Contains(uri.Host) ? Task.FromResult(token) : Task.FromResult(string.Empty);
    }
}
