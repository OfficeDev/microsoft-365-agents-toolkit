using {{YOUR_NAMESPACE}}.Configuration;
using Microsoft.Agents.Builder;
using Microsoft.Graph;
using Microsoft.Kiota.Abstractions.Authentication;

namespace {{YOUR_NAMESPACE}}.SSO;

public class TokenProvider : IAccessTokenProvider
{
  private string token { get; set; }

  public TokenProvider(String token)
  {
    this.token = token;
  }

  public Task<string> GetAuthorizationTokenAsync(Uri uri, Dictionary<string, object> additionalAuthenticationContext = default,
      CancellationToken cancellationToken = default)
  {
    // get the token and return it
    return Task.FromResult(this.token);
  }

  public AllowedHostsValidator AllowedHostsValidator { get; }
}

public static class SsoOperations
{
    public static async Task ShowUserInfo(ITurnContext stepContext, string token, BotAuthenticationOptions botAuthOptions)
    {
        await stepContext.SendActivityAsync("Retrieving user information from Microsoft Graph ...");
        var tokenCredential = new BaseBearerTokenAuthenticationProvider(new TokenProvider(token));
        var graphClient = new GraphServiceClient(tokenCredential);
        var profile = await graphClient.Me.GetAsync();
        await stepContext.SendActivityAsync($"You're logged in as {profile.DisplayName}");
    }
}