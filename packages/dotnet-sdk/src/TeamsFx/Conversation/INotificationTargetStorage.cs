namespace Microsoft.TeamsFx.Conversation
{
    using Microsoft.Agents.Core.Models;

    /// <summary>
    /// Interface for a storage provider that stores and retrieves notification target references.
    /// </summary>
    [Obsolete($"Use {nameof(IConversationReferenceStore)} to customize the way to persist bot notification connections instead.")]
    public interface INotificationTargetStorage
    {
        /// <summary>
        /// Read one notification target by its key.
        /// </summary>
        /// <param name="key">The target key</param>
        /// <param name="cancellationToken">The cancellation token.</param>
        /// <returns>The target, or null if not found.</returns>
        Task<ConversationReference> Read(string key, CancellationToken cancellationToken = default);

        /// <summary>
        /// List all stored notification targets.
        /// </summary>
        /// <param name="cancellationToken">The cancellation token.</param>
        /// <returns>An array of notification target. Or an empty array if nothing is stored.</returns>
        Task<ConversationReference[]> List(CancellationToken cancellationToken = default);

        /// <summary>
        /// Write one notification target by its key.
        /// </summary>
        /// <param name="key">The target key.</param>
        /// <param name="reference">The target object.</param>
        /// <param name="cancellationToken">The cancellation token.</param>
        /// <returns>The task object representing the asynchronous operation.</returns>
        Task Write(string key, ConversationReference reference, CancellationToken cancellationToken = default);

        /// <summary>
        /// Delete one notification target by its key.
        /// </summary>
        /// <param name="key">The target key.</param>
        /// <param name="cancellationToken">The cancellation token.</param>
        /// <returns>The task object representing the asynchronous operation.</returns>
        Task Delete(string key, CancellationToken cancellationToken = default);
    }
}
