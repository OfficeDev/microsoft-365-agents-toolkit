// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

namespace Microsoft.TeamsFx.Conversation
{
    /// <summary>
    /// The options to add a conversation reference.
    /// </summary>
    [Obsolete("This package will be deprecated by 2026-09. Please use Microsoft 365 Agents SDK (https://github.com/microsoft/Agents-for-net) instead.")]
    public class ConversationReferenceStoreAddOptions
    {
        /// <summary>
        /// Gets or sets a value indicating whether to overwrite the existing conversation reference.
        /// </summary>
        public bool? Overwrite { get; set; }
    }
}
