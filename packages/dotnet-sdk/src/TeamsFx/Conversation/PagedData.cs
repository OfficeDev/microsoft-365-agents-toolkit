// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

namespace Microsoft.TeamsFx.Conversation
{
    /// <summary>
    /// Represents a page of data.
    /// </summary>
    [Obsolete("This package will be deprecated by 2026-09. Please use Microsoft 365 Agents SDK (https://www.nuget.org/packages/Microsoft.Agents.Hosting.AspNetCore) instead.")]
    public class PagedData<T>
    {
        /// <summary>
        /// Gets or sets the page of items.
        /// </summary>
        /// <value>
        /// The array of items.
        /// </value>
        public T[] Data { get; set; } = Array.Empty<T>();

        /// <summary>
        /// Gets or sets a token for retrieving the next page of results.
        /// </summary>
        /// <value>
        /// The Continuation Token to pass to get the next page of results.
        /// Null or empty token means the page reaches the end.
        /// </value>
        public string ContinuationToken { get; set; }
    }
}
