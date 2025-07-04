using Microsoft.Azure.Functions.Worker;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.Agents.Builder;
using Microsoft.Agents.Hosting.AspNetCore;
using {{SafeProjectName}}.Notification;

namespace {{SafeProjectName}}
{
    public sealed class MessageHandler
    {
        private readonly NotificationBot _notification;
        private readonly IAgent _bot;
        private readonly ILogger<MessageHandler> _log;

        public MessageHandler(NotificationBot notification, IAgent bot, ILogger<MessageHandler> log)
        {
            _notification = notification;
            _bot = bot;
            _log = log;
        }

        [Function("MessageHandler")]
        public async Task<EmptyResult> Run([HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = "api/messages")] HttpRequest req)
        {
            _log.LogInformation("MessageHandler processes a request.");

            await (_notification.Adapter as CloudAdapter).ProcessAsync(req, req.HttpContext.Response, _bot, req.HttpContext.RequestAborted);

            return new EmptyResult();
        }
    }
}
