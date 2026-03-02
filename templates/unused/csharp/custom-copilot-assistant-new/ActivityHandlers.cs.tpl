using Microsoft.Bot.Builder;
using Microsoft.Teams.AI;
using {{SafeProjectName}}.Model;

namespace {{SafeProjectName}}
{
    /// <summary>
    /// Defines the activity handlers.
    /// </summary>
    public static class ActivityHandlers
    {
        /// <summary>
        /// Handles "/reset" message.
        /// </summary>
        public static RouteHandler<AppState> ResetMessageHandler = async (ITurnContext turnContext, AppState turnState, CancellationToken cancellationToken) =>
        {
            turnState.DeleteConversationState();
            await turnContext.SendActivityAsync("Ok lets start this over.", cancellationToken: cancellationToken);
        };
    }
}
