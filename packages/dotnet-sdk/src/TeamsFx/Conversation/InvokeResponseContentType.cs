// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

namespace Microsoft.TeamsFx.Conversation
{
    [Obsolete("This package will be deprecated by 2026-09. Please use Microsoft 365 Agents SDK (https://www.nuget.org/packages/Microsoft.Agents.Hosting.AspNetCore) instead.")]
    internal static class InvokeResponseContentType
    {
        public const string AdaptiveCard = "application/vnd.microsoft.card.adaptive";
        public const string Message = "application/vnd.microsoft.activity.message";
        public const string Error = "application/vnd.microsoft.error";
    }
}
