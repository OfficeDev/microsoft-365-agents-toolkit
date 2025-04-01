// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

namespace Microsoft.TeamsFx.Conversation
{
    using Microsoft.Agents.Core.Models;

    [Obsolete]
    internal sealed class DefaultConversationReferenceStore : IConversationReferenceStore
    {
        private readonly INotificationTargetStorage _storage;

        public DefaultConversationReferenceStore(INotificationTargetStorage storage)
        {
            _storage = storage;
        }

        public async Task<bool> Add(
            string key,
            ConversationReference reference,
            ConversationReferenceStoreAddOptions options,
            CancellationToken cancellationToken = default)
        {
            if (string.IsNullOrEmpty(key))
            {
                throw new ArgumentException($"{nameof(key)} can't be null or empty.");
            }

            if (reference == null)
            {
                throw new ArgumentNullException(nameof(reference));
            }

            if (options == null)
            {
                throw new ArgumentNullException(nameof(options));
            }

            var overwrite = options.Overwrite ?? false;
            if (overwrite)
            {
                await _storage.Write(key, reference, cancellationToken).ConfigureAwait(false);
                return true;
            }

            var result = await _storage.Read(key, cancellationToken).ConfigureAwait(false);
            if (result == null)
            {
                await _storage.Write(key, reference, cancellationToken).ConfigureAwait(false);
                return true;
            }

            return false;
        }

        public async Task<bool> Remove(
            string key,
            ConversationReference reference,
            CancellationToken cancellationToken = default)
        {
            if (string.IsNullOrEmpty(key))
            {
                throw new ArgumentException($"{nameof(key)} can't be null or empty.");
            }

            if (reference == null)
            {
                throw new ArgumentNullException(nameof(reference));
            }

            var result = await _storage.Read(key, cancellationToken).ConfigureAwait(false);
            if (result == null)
            {
                return false;
            }

            await _storage.Delete(key, cancellationToken).ConfigureAwait(false);
            return true;
        }

        public async Task<PagedData<ConversationReference>> List(
            int? pageSize = null,
            string continuationToken = null,
            CancellationToken cancellationToken = default)
        {
            var data = await _storage.List(cancellationToken).ConfigureAwait(false);
            return new PagedData<ConversationReference> {
                Data = data,
                ContinuationToken = null,
            };
        }
    }
}
