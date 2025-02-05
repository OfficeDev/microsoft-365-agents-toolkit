using {{SafeProjectName}}.Models;
using AdaptiveCards.Templating;
using Microsoft.Azure.Functions.Worker;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.TeamsFx.Conversation;
using Newtonsoft.Json;

namespace {{SafeProjectName}}
{
    public sealed class NotifyHttpTrigger
    {
        private readonly ConversationBot _conversation;
        private readonly ILogger<NotifyHttpTrigger> _log;
        private readonly string _contentRootPath;

        public NotifyHttpTrigger(ConversationBot conversation, ILogger<NotifyHttpTrigger> log, string contentRootPath)
        {
            _conversation = conversation;
            _log = log;
            _contentRootPath = contentRootPath;
        }

        [Function("NotifyHttpTrigger")]
        public async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = "api/notification")] HttpRequest req, ExecutionContext context)
        {
            _log.LogInformation("NotifyHttpTrigger is triggered.");

            // Read adaptive card template
            var adaptiveCardFilePath = Path.Combine(_contentRootPath, "Resources", "NotificationDefault.json");
            var cardTemplate = await File.ReadAllTextAsync(adaptiveCardFilePath, req.HttpContext.RequestAborted);

            var pageSize = 100;
            string continuationToken = null;
            do
            {
                var pagedInstallations = await _conversation.Notification.GetPagedInstallationsAsync(pageSize, continuationToken, req.HttpContext.RequestAborted);
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
                            Description = $"This is a sample http-triggered notification to {installation.Type}",
                            NotificationUrl = "https://aka.ms/teamsfx-notification-new",
                        }
                    );
                    await installation.SendAdaptiveCard(JsonConvert.DeserializeObject(cardContent), req.HttpContext.RequestAborted);
                }

            } while (!string.IsNullOrEmpty(continuationToken));

            return new OkResult();
        }
    }
}
