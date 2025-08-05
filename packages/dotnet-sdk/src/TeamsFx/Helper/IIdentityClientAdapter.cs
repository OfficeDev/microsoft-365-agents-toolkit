// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using Microsoft.Identity.Client;

namespace Microsoft.TeamsFx.Helper
{
    /// <summary>
    /// Adapter of IConfidentialClientApplication On-behalf-of flow.
    /// </summary>
    [Obsolete("This package will be deprecated by 2026-09. Please use Microsoft 365 Agents SDK (https://github.com/microsoft/Agents-for-net) instead.")]
    public interface IIdentityClientAdapter
    {
        /// <summary>
        /// Use On-behalf-of flow to exchange access token.
        /// </summary>
        /// <param name="ssoToken">token from Teams client</param>
        /// <param name="scopes">required scopes</param>
        /// <returns></returns>
        [Obsolete("This package will be deprecated by 2026-09. Please use Microsoft 365 Agents SDK (https://github.com/microsoft/Agents-for-net) instead.")]
        Task<AuthenticationResult> GetAccessToken(string ssoToken, IEnumerable<string> scopes);
    }
}
