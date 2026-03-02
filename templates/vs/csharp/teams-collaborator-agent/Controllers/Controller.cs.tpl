using System.Text.Json;
using {{SafeProjectName}}.Agent;
using {{SafeProjectName}}.Capability;
using {{SafeProjectName}}.Models;
using {{SafeProjectName}}.Storage;
using {{SafeProjectName}}.Utils;
using Microsoft.Teams.Api.Activities;
using Microsoft.Teams.Api.Activities.Invokes;
using Microsoft.Teams.Api.Entities;
using Microsoft.Teams.Apps;
using Microsoft.Teams.Apps.Activities;
using Microsoft.Teams.Apps.Activities.Invokes;
using Microsoft.Teams.Apps.Annotations;
using MessageActivity = Microsoft.Teams.Api.Activities.MessageActivity;

namespace {{SafeProjectName}}.Controllers
{
    [TeamsController]
    public class Controller(
        ILogger<Controller> logger,
        ILoggerFactory loggerFactory,
        IDatabase database,
        List<CapabilityDefinition> capabilityDefinitions
    )
    {
        [Message]
        public async Task OnMessage(IContext<MessageActivity> context)
        {
            var activity = context.Activity;

            var botMentioned = activity.Entities?.Any(e => e is MentionEntity) ?? false;
            var shouldProcess = activity.Conversation.IsGroup != true || botMentioned;

            var messageContext = await MessageContextFactory.CreateAsync(
                context,
                database,
                logger: logger,
                includeMembers: shouldProcess && activity.Conversation.IsGroup == true
            );

            var trackedMessages = MessageUtils.CreateMessageRecords(new[] { activity });

            if (shouldProcess)
            {
                var manager = new Manager(loggerFactory, capabilityDefinitions);
                var result = await manager.ProcessRequestAsync(messageContext);

                var finalizedMessage = MessageUtils.FinalizePromptResponse(
                    result.Response,
                    messageContext,
                    logger
                );

                var response = await context.Send(finalizedMessage);
                trackedMessages = MessageUtils.CreateMessageRecords(new[] { activity, response });
            }

            await messageContext.Memory.AddMessagesAsync(trackedMessages);
            logger.LogDebug(
                "Saved {Count} messages to conversation {ConversationId}",
                trackedMessages.Count,
                messageContext.ConversationId
            );
        }

        [Microsoft.Teams.Apps.Activities.Invokes.Message.SubmitAction]
        public async Task OnSubmitAction(IContext<Messages.SubmitActionActivity> context)
        {
            try
            {
                var actionValue = context.Activity.Value.ActionValue;

                string reaction = "unknown";
                object? feedbackJson = null;

                if (actionValue != null)
                {
                    try
                    {
                        var json = JsonSerializer.Serialize(actionValue);
                        var dict = JsonSerializer.Deserialize<Dictionary<string, object>>(json);

                        if (dict != null)
                        {
                            if (dict.TryGetValue("reaction", out var reactionObj))
                            {
                                reaction = reactionObj?.ToString() ?? "unknown";
                            }

                            if (dict.TryGetValue("feedback", out var feedbackObj))
                            {
                                feedbackJson = feedbackObj;
                            }
                        }
                    }
                    catch
                    {
                        logger.LogWarning("Failed to parse submit action value payload");
                    }
                }

                if (string.IsNullOrEmpty(context.Activity.ReplyToId))
                {
                    logger.LogWarning(
                        "No replyToId found for messageId {MessageId}",
                        context.Activity.Id
                    );
                    return;
                }

                var success = await database.RecordFeedbackAsync(
                    context.Activity.ReplyToId,
                    reaction,
                    feedbackJson
                );

                if (success)
                {
                    logger.LogDebug(
                        "Recorded feedback for message {ReplyToId}",
                        context.Activity.ReplyToId
                    );
                }
                else
                {
                    logger.LogWarning(
                        "Failed to record feedback for message {ReplyToId}",
                        context.Activity.ReplyToId
                    );
                }
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Error processing feedback: {ErrorMessage}", ex.Message);
            }
        }

        [Conversation.MembersAdded]
        public async Task OnMembersAdded(IContext<ConversationUpdateActivity> context)
        {
            foreach (var member in context.Activity.MembersAdded)
            {
                if (member.Id != context.Activity.Recipient.Id)
                {
                    await context.Send(
                        "Hi! I'm the Collab Agent. I'll listen to the conversation and can provide summaries, action items, or search for a message when asked."
                    );
                }
            }
        }
    }
}
