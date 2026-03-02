// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using Microsoft.Identity.Client;

namespace {{YOUR_NAMESPACE}}.SSO
{
  /// <summary>
  /// Adapter of IConfidentialClientApplication On-behalf-of flow.
  /// </summary>
  public interface IIdentityClientAdapter
  {
    /// <summary>
    /// Use On-behalf-of flow to exchange access token.
    /// </summary>
    /// <param name="ssoToken">token from Teams client</param>
    /// <param name="scopes">required scopes</param>
    /// <returns></returns>
    Task<AuthenticationResult> GetAccessToken(string ssoToken, IEnumerable<string> scopes);
  }
}
