namespace {{SafeProjectName}}.Notification
{
    using Microsoft.Agents.Core.Models;
    using Microsoft.Agents.Storage;

    public interface IConversationReferenceStorage: IStorage
    {
        Task<PagedData<ConversationReference>> ListAsync(int? pageSize = null, string continuationToken = null, CancellationToken cancellationToken = default);
    }
}
