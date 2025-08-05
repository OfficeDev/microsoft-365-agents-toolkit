// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using Microsoft.Agents.Core.Models;
namespace {{YOUR_NAMESPACE}}.SSO;

/// <summary>
/// Token response provided by Teams Bot SSO prompt
/// </summary>
public class TeamsBotSsoPromptTokenResponse : TokenResponse
{
  /// <summary>
  /// SSO token for user
  /// </summary>
  public string SsoToken { get; set; }

  /// <summary>
  /// Expiration time of SSO token
  /// </summary>
  public string SsoTokenExpiration { get; set; }
}
