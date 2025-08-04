// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using Azure.Core;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.Identity.Client;
using Microsoft.JSInterop;
using Microsoft.TeamsFx.Configuration;
using Microsoft.TeamsFx.Helper;
using System.Net;

using AccessToken = Microsoft.TeamsFx.Model.AccessToken;

namespace Microsoft.TeamsFx;

/// <summary>
/// Represent Teams current user's identity, and it is used within Teams tab application.
/// </summary>
/// <remarks>
/// Can only be used within Blazor server for security reason.
/// </remarks>
[Obsolete("This package will be deprecated by 2026-09. Please use Microsoft 365 Agents SDK (https://www.nuget.org/packages/Microsoft.Agents.Core) instead.")]
public class TeamsUserCredential : TokenCredential, IAsyncDisposable
{
    internal bool _initialized;
    internal AccessToken _ssoToken;
    private readonly AuthenticationOptions _authenticationOptions;
    private IIdentityClientAdapter _identityClientAdapter;

    #region JS Interop
    private readonly Lazy<Task<IJSObjectReference>> _teamsSdkTask;
    private readonly IJSRuntime _jsRuntime;
    #endregion

    #region Util
    private readonly ILogger<TeamsUserCredential> _logger;
    #endregion

    /// <summary>
    /// Constructor of TeamsUserCredential.
    /// Developer need to configure TeamsFx service before using this class.
    /// </summary>
    /// <param name="authenticationOptions">Authentication options filled by DI.</param>
    /// <param name="jsRuntime">JavaScript interop runtime.</param>
    /// <param name="logger">Logger of TeamsUserCredential Class.</param>
    /// <param name="identityClientAdapter">Global instance of adaptor to call MSAL.NET library</param>
    /// <exception cref="ExceptionCode.InvalidConfiguration">When client id, client secret, initiate login endpoint or OAuth authority is missing or invalid in config.</exception>
    [Obsolete("This package will be deprecated by 2026-09. Please use Microsoft 365 Agents SDK (https://www.nuget.org/packages/Microsoft.Agents.Core) instead.")]
    public TeamsUserCredential(
        IOptions<AuthenticationOptions> authenticationOptions,
        IJSRuntime jsRuntime,
        ILogger<TeamsUserCredential> logger,
        IIdentityClientAdapter identityClientAdapter)
    {
        _logger = logger;
        try
        {
            _logger.LogTrace("Validate authentication configuration");
            _authenticationOptions = authenticationOptions.Value;
        }
        catch (OptionsValidationException e)
        {
            throw new ExceptionWithCode($"Authentication config is missing or not correct with error: {e.Message}", ExceptionCode.InvalidConfiguration);
        }
        _identityClientAdapter = identityClientAdapter;
        _teamsSdkTask = new(() => ImportTeamsSdk(jsRuntime).AsTask());
        _jsRuntime = jsRuntime;

        logger.LogInformation("Create teams user credential");
    }

    /// <summary>
    /// Get basic information of current Teams user.
    /// <example>
    /// For example:
    /// <code>
    /// var user = await teamsUserCredential.GetUserInfoAsync();
    /// </code>
    /// </example>
    /// </summary>
    /// <returns>Basic user info with user displayName, objectId and preferredUserName.</returns>
    /// <exception cref="ExceptionCode.InternalError">When SSO token from Teams client is not valid.</exception>
    /// <exception cref="ExceptionCode.InvalidParameter">When SSO token from Teams client is empty.</exception>
    [Obsolete("This package will be deprecated by 2026-09. Please use Microsoft 365 Agents SDK (https://www.nuget.org/packages/Microsoft.Agents.Core) instead.")]
    public async ValueTask<UserInfo> GetUserInfoAsync()
    {
        _logger.LogInformation("Get basic user info from SSO token");

        await EnsureTeamsSdkInitialized().ConfigureAwait(false);
        var ssoToken = await GetSsoTokenAsync().ConfigureAwait(false);
        try
        {
            var userInfo = Utils.GetUserInfoFromSsoToken(ssoToken.Token);
            return userInfo;
        }
        catch (ExceptionWithCode e)
        {
            _logger.LogError(e.Message);
            throw;
        }
    }

    /// <summary>
    /// Popup login page to get user's access token with specific scopes.
    /// </summary>
    ///
    /// <param name="scopes">The string of Microsoft Token scopes of access. Default value is `.default`.</param>
    ///
    /// <example>
    /// For example:
    /// <code>
    /// await teamsUserCredential.LoginAsync("User.Read"); // single scopes using string
    /// await teamsUserCredential.LoginAsync("User.Read Calendars.Read"); // multiple scopes using string
    /// </code>
    /// </example>
    ///
    /// <remarks>
    /// Only works in Teams client app. User will be redirected to the authorization page to login and consent.
    /// </remarks>
    ///
    /// <exception cref="ExceptionCode.InternalError">When failed to login with unknown error.</exception>
    /// <exception cref="ExceptionCode.ConsentFailed">When user canceled or failed to consent.</exception>
    [Obsolete("This package will be deprecated by 2026-09. Please use Microsoft 365 Agents SDK (https://www.nuget.org/packages/Microsoft.Agents.Core) instead.")]
    public async Task LoginAsync(string scopes)
    {
        _logger.LogInformation($"Popup consent page to get user's access token with scopes: {scopes}");

        await EnsureTeamsSdkInitialized().ConfigureAwait(false);
        try
        {
            var teamsSdk = await _teamsSdkTask.Value.ConfigureAwait(false);
            var url = $"{_authenticationOptions.InitiateLoginEndpoint}?clientId={_authenticationOptions.ClientId}&scope={WebUtility.UrlEncode(scopes)}";
            var token = await teamsSdk.InvokeAsync<string>("authenticate", url).ConfigureAwait(false);
            if (string.IsNullOrEmpty(token))
            {
                var errorMessage = "Get empty authentication result from Teams";
                _logger.LogError(errorMessage);
                throw new ExceptionWithCode(errorMessage, ExceptionCode.InternalError);
            }
        }
        catch (JSException e)
        {
            var errorMessage = $"Consent failed for the scope {scopes} with error: {e.Message}";
            _logger.LogError(errorMessage);
            throw new ExceptionWithCode(errorMessage, ExceptionCode.ConsentFailed);
        }
    }

    /// <summary>
    /// Popup login page to get user's access token with specific scopes.
    /// </summary>
    ///
    /// <param name="scopes">The array of Microsoft Token scopes of access. Default value is `[.default]`.</param>
    ///
    /// <example>
    /// For example:
    /// <code>
    /// await teamsUserCredential.LoginAsync(new string[] { "User.Read" }); // single scope using string array
    /// await teamsUserCredential.LoginAsync(new string[] { "User.Read Calendars.Read" }); //  multiple scopes using string array
    /// </code>
    /// </example>
    ///
    /// <remarks>
    /// Only works in Teams client APP. User will be redirected to the authorization page to login and consent.
    /// </remarks>
    ///
    /// <exception cref="ExceptionCode.InternalError">When failed to login with unknown error.</exception>
    /// <exception cref="ExceptionCode.ConsentFailed">When user canceled or failed to consent.</exception>
    [Obsolete("This package will be deprecated by 2026-09. Please use Microsoft 365 Agents SDK (https://www.nuget.org/packages/Microsoft.Agents.Core) instead.")]
    public async Task LoginAsync(string[] scopes)
    {
        var scopeString = string.Join(' ', scopes);
        await LoginAsync(scopeString).ConfigureAwait(false);
    }

    /// <summary>
    /// Gets an Azure.Core.AccessToken for the specified set of scopes.
    /// </summary>
    /// <param name="requestContext">The Azure.Core.TokenRequestContext with authentication information.</param>
    /// <param name="cancellationToken">The System.Threading.CancellationToken to use.</param>
    /// <returns>A valid Azure.Core.AccessToken.</returns>
    ///
    /// <exception cref="ExceptionCode.InternalError">When failed to login with unknown error.</exception>
    /// <exception cref="ExceptionCode.UiRequiredError">When need user consent to get access token.</exception>
    /// <exception cref="ExceptionCode.ServiceError">When failed to get access token from identity server(AAD).</exception>
    ///
    /// <example>
    /// For example:
    /// <code>
    /// // Get Graph access token for single scope
    /// await teamsUserCredential.GetTokenAsync(new TokenRequestContext(new string[] { "User.Read" }), new CancellationToken());
    /// // Get Graph access token for multiple scopes
    /// await teamsUserCredential.GetTokenAsync(new TokenRequestContext(new string[] { "User.Read", "Application.Read.All" }), new CancellationToken());
    /// </code>
    /// </example>
    /// <remarks>
    /// Can only be used within Teams.
    /// </remarks>
    [Obsolete("This package will be deprecated by 2026-09. Please use Microsoft 365 Agents SDK (https://www.nuget.org/packages/Microsoft.Agents.Core) instead.")]
    public async override ValueTask<global::Azure.Core.AccessToken> GetTokenAsync(TokenRequestContext requestContext, CancellationToken cancellationToken)
    {
        await EnsureTeamsSdkInitialized().ConfigureAwait(false);
        var scopes = requestContext.Scopes;
        var ssoToken = await GetSsoTokenAsync().ConfigureAwait(false);
        if (scopes == null || scopes.Length == 0)
        {
            _logger.LogInformation("Get SSO token");
            return ssoToken.ToAzureAccessToken();
        }
        else
        {
            _logger.LogInformation($"Get access token with scopes: {string.Join(' ', scopes)}");
            var accessToken = await GetAccessTokenByOnBehalfOfFlow(ssoToken.Token, scopes).ConfigureAwait(false);
            return accessToken;
        }
    }

    /// <summary>
    /// Sync version is not supported now. Please use GetTokenAsync instead.
    /// </summary>
    [Obsolete("This package will be deprecated by 2026-09. Please use Microsoft 365 Agents SDK (https://www.nuget.org/packages/Microsoft.Agents.Core) instead.")]
    public override global::Azure.Core.AccessToken GetToken(TokenRequestContext requestContext, CancellationToken cancellationToken)
    {
        throw new NotImplementedException();
    }

    /// <summary>
    /// Dispose.
    /// </summary>
    /// <returns></returns>
    [Obsolete("This package will be deprecated by 2026-09. Please use Microsoft 365 Agents SDK (https://www.nuget.org/packages/Microsoft.Agents.Core) instead.")]
    public async ValueTask DisposeAsync()
    {
        if (_teamsSdkTask.IsValueCreated)
        {
            var instance = await _teamsSdkTask.Value.ConfigureAwait(false);
            await instance.DisposeAsync().ConfigureAwait(false);
        }
    }

    private async ValueTask<IJSObjectReference> ImportTeamsSdk(IJSRuntime jsRuntime)
    {
        try
        {
            await jsRuntime.InvokeVoidAsync("import", "https://res.cdn.office.net/teams-js/2.22.0/js/MicrosoftTeams.min.js").ConfigureAwait(false);
            return await jsRuntime.InvokeAsync<IJSObjectReference>("import", "./_content/Microsoft.TeamsFx/jsInterop.js").AsTask().ConfigureAwait(false);
        }
        catch (JSException e)
        {
            throw new ExceptionWithCode(e.Message, ExceptionCode.JSRuntimeError);
        }
    }

    private async Task EnsureTeamsSdkInitialized()
    {
        if (_initialized)
        {
            return;
        }
        try
        {
            var teamsSdk = await _teamsSdkTask.Value.ConfigureAwait(false);
            await teamsSdk.InvokeVoidAsync("initialize").ConfigureAwait(false);
            _initialized = true;
        }
        catch (JSException e)
        {
            if (e.Message == "timeout")
            {
                var errorMsg = "Initialize teams sdk timeout, maybe the code is not running inside Teams";
                _logger.LogError(errorMsg);
                throw new ExceptionWithCode(errorMsg, ExceptionCode.InternalError);
            }
        }
    }

    /// <summary>
    /// Get SSO token using teams SDK.
    /// It will try to get SSO token from memory first, if SSO token doesn't exist or about to expired, then it will using teams SDK to get SSO token.
    /// </summary>
    /// <returns></returns>
    private async Task<AccessToken> GetSsoTokenAsync()
    {
        if (_ssoToken != null)
        {
            if (!_ssoToken.NearExpiration())
            {
                _logger.LogTrace("Get SSO token from memory cache");
                return _ssoToken;
            }
        }

        string token;
        try
        {
            var teamsSdk = await _teamsSdkTask.Value.ConfigureAwait(false);
            token = await teamsSdk.InvokeAsync<string>("getAuthToken").ConfigureAwait(false);
            if (string.IsNullOrEmpty(token))
            {
                var errorMessage = "Get empty SSO token from Teams";
                _logger.LogError(errorMessage);
                throw new ExceptionWithCode(errorMessage, ExceptionCode.InternalError);
            }
        }
        catch (JSException e)
        {
            var errorMessage = $"Get SSO token failed with error: {e.Message}";
            _logger.LogError(errorMessage);
            throw new ExceptionWithCode(errorMessage, ExceptionCode.InternalError);
        }
        // Validate token version
        try
        {
            var tokenObject = Utils.ParseJwt(token);
            var version = tokenObject.Payload["ver"].ToString();
            if (version != "1.0" && version != "2.0")
            {
                var errorMessage = $"SSO token is not valid with an unknown version: {version}";
                _logger.LogError(errorMessage);
                throw new ExceptionWithCode(errorMessage, ExceptionCode.InternalError);
            }

            var exp = tokenObject.Payload["exp"].ToString();
            var expireTime = DateTimeOffset.FromUnixTimeSeconds(long.Parse(exp));
            _ssoToken = new AccessToken(token, expireTime);
            return _ssoToken;
        }
        catch (ExceptionWithCode e)
        {
            _logger.LogError(e.Message);
            throw;
        }
    }

    /// <summary>
    /// Get access token from identity server (AAD).
    /// </summary>
    /// <param name="ssoToken">token returned from Teams SDK</param>
    /// <param name="scopes">required scopes</param>
    /// <returns></returns>
    private async ValueTask<global::Azure.Core.AccessToken> GetAccessTokenByOnBehalfOfFlow(string ssoToken, IEnumerable<string> scopes)
    {
        _logger.LogTrace($"Get access token from authentication server with scopes: {string.Join(' ', scopes)}");

        try
        {
            _logger.LogDebug("Acquiring token via OBO flow.");
            var result = await _identityClientAdapter
                .GetAccessToken(ssoToken, scopes)
                .ConfigureAwait(false);

            var accessToken = new global::Azure.Core.AccessToken(result.AccessToken, result.ExpiresOn);
            return accessToken;
        }
        catch (MsalUiRequiredException) // Need user interaction
        {
            var fullErrorMsg = $"Failed to get access token from OAuth identity server, please login(consent) first";
            _logger.LogWarning(fullErrorMsg);
            throw new ExceptionWithCode(fullErrorMsg, ExceptionCode.UiRequiredError);
        }
        catch (MsalServiceException ex) // Errors that returned from AAD service
        {
            var fullErrorMsg = $"Failed to get access token from OAuth identity server with error: {ex.ResponseBody}";
            _logger.LogWarning(fullErrorMsg);
            throw new ExceptionWithCode(fullErrorMsg, ExceptionCode.ServiceError);
        }
        catch (MsalClientException ex) // Exceptions that are local to the MSAL library
        {
            var fullErrorMsg = $"Failed to get access token with error: {ex.Message}";
            _logger.LogWarning(fullErrorMsg);
            throw new ExceptionWithCode(fullErrorMsg, ExceptionCode.InternalError);
        }
    }
}
