using Microsoft.Bot.Builder;
using Microsoft.Teams.AI;
using Microsoft.Teams.AI.AI.Planners.Experimental;

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
        public static RouteHandler<AssistantsState> ResetMessageHandler = async (ITurnContext turnContext, AssistantsState turnState, CancellationToken cancellationToken) =>
        {
            turnState.DeleteConversationState();
            await turnContext.SendActivityAsync("Ok lets start this over.", cancellationToken: cancellationToken);
        };
    }
}
