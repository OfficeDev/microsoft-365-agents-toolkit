// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

namespace Microsoft.TeamsFx.Conversation
{
    /// <summary>
    /// Status code for an application/vnd.microsoft.error invoke response.
    /// </summary>
    [Obsolete("This package will be deprecated by 2026-09. Please use Microsoft 365 Agents SDK (https://www.nuget.org/packages/Microsoft.Agents.Core) instead.")]
    public enum InvokeResponseErrorCode
    {
        /// <summary>
        /// Invalid request.
        /// </summary>
        BadRequest = 400,

        /// <summary>
        /// Internal server error.
        /// </summary>
        InternalServerError = 500,
    }
}
