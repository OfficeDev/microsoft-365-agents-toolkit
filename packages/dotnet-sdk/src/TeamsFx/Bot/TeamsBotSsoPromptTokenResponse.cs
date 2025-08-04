// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using Microsoft.Agents.Core.Models;
namespace Microsoft.TeamsFx.Bot;

/// <summary>
/// Token response provided by Teams Bot SSO prompt
/// </summary>
[Obsolete("This package will be deprecated by 2026-09. Please use Microsoft 365 Agents SDK (https://www.nuget.org/packages/Microsoft.Agents.Core) instead.")]
public class TeamsBotSsoPromptTokenResponse: TokenResponse
{
    /// <summary>
    /// SSO token for user
    /// </summary>
    [Obsolete("This package will be deprecated by 2026-09. Please use Microsoft 365 Agents SDK (https://www.nuget.org/packages/Microsoft.Agents.Core) instead.")]
    public string SsoToken { get; set; }

    /// <summary>
    /// Expiration time of SSO token
    /// </summary>
    [Obsolete("This package will be deprecated by 2026-09. Please use Microsoft 365 Agents SDK (https://www.nuget.org/packages/Microsoft.Agents.Core) instead.")]
    public string SsoTokenExpiration { get; set; }
}
