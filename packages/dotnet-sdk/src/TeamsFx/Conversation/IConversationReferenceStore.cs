// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

namespace Microsoft.TeamsFx.Conversation
{
    using Microsoft.Agents.Core.Models;

    /// <summary>
    /// Interface for a store provider that manages notification target references.
    /// </summary>
    public interface IConversationReferenceStore
    {
        /// <summary>
        /// Add a conversation reference to the store. If overwrite, update the existing one, otherwise add when not exist.
        /// </summary>
        /// <param name="key">The target key.</param>
        /// <param name="reference">The conversation reference to be added.</param>
        /// <param name="options">The options to add the conversation reference.</param>
        /// <param name="cancellationToken">The cancellation token.</param>
        /// <returns>true if added or updated, false if not changed.</returns>
        public Task<bool> Add(
            string key,
            ConversationReference reference,
            ConversationReferenceStoreAddOptions options,
            CancellationToken cancellationToken = default);

        /// <summary>
        /// Remove a conversation reference from the store.
        /// </summary>
        /// <param name="key">The target key.</param>
        /// <param name="reference">The conversation reference to be removed.</param>
        /// <param name="cancellationToken">The cancellation token.</param>
        /// <returns>true if exist and removed, false if not changed.</returns>
        public Task<bool> Remove(
            string key,
            ConversationReference reference,
            CancellationToken cancellationToken = default);


        /// <summary>
        /// List stored conversation reference by page.
        /// </summary>
        /// <param name="pageSize">The page size.</param>
        /// <param name="continuationToken">The continuation token.</param>
        /// <param name="cancellationToken">The cancellation token.</param>
        /// <returns>A paged list of conversation references.</returns>
        public Task<PagedData<ConversationReference>> List(
            int? pageSize = default,
            string continuationToken = default,
            CancellationToken cancellationToken = default);
    }
}
