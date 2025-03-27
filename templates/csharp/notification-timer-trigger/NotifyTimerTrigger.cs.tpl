using {{SafeProjectName}}.Models;
using AdaptiveCards.Templating;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Microsoft.TeamsFx.Conversation;

namespace {{SafeProjectName}}
{
    public sealed class NotifyTimerTrigger
    {
        private readonly ConversationBot _conversation;
        private readonly ILogger<NotifyTimerTrigger> _log;

        public NotifyTimerTrigger(ConversationBot conversation, ILogger<NotifyTimerTrigger> log)
        {
            _conversation = conversation;
            _log = log;
        }

        [Function("NotifyTimerTrigger")]
        public async Task Run([TimerTrigger("*/30 * * * * *")]TimerInfo myTimer, ExecutionContext context, CancellationToken cancellationToken)
        {
            _log.LogInformation($"NotifyTimerTrigger is triggered at {DateTime.Now}.");

            // Read adaptive card template
            var adaptiveCardFilePath = Path.Combine("Resources", "NotificationDefault.json");
            var cardTemplate = await File.ReadAllTextAsync(adaptiveCardFilePath, cancellationToken);

            var pageSize = 100;
            string continuationToken = null;
            do
            {
                var pagedInstallations = await _conversation.Notification.GetPagedInstallationsAsync(pageSize, continuationToken, cancellationToken);
                continuationToken = pagedInstallations.ContinuationToken;
                var installations = pagedInstallations.Data;
                foreach (var installation in installations)
                {
                    // Build and send adaptive card
                    var cardContent = new AdaptiveCardTemplate(cardTemplate).Expand
                    (
                        new NotificationDefaultModel
                        {
                            Title = "New Event Occurred!",
                            AppName = "Contoso App Notification",
                            Description = $"This is a sample timer-triggered notification to {installation.Type}",
                            NotificationUrl = "https://aka.ms/teamsfx-notification-new",
                        }
                    );
                    await installation.SendAdaptiveCard(cardContent, cancellationToken);
                }

            } while (!string.IsNullOrEmpty(continuationToken));
        }
    }
}
