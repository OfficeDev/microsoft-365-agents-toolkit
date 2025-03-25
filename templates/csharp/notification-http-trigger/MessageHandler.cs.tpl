using Microsoft.Azure.Functions.Worker;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.TeamsFx.Conversation;
using Microsoft.Agents.BotBuilder;
using Microsoft.Agents.Hosting.AspNetCore;

namespace {{SafeProjectName}}
{
    public sealed class MessageHandler
    {
        private readonly ConversationBot _conversation;
        private readonly IBot _bot;
        private readonly ILogger<MessageHandler> _log;

        public MessageHandler(ConversationBot conversation, IBot bot, ILogger<MessageHandler> log)
        {
            _conversation = conversation;
            _bot = bot;
            _log = log;
        }

        [Function("MessageHandler")]
        public async Task<EmptyResult> Run([HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = "api/messages")] HttpRequest req)
        {
            _log.LogInformation("MessageHandler processes a request.");

            await (_conversation.Adapter as CloudAdapter).ProcessAsync(req, req.HttpContext.Response, _bot, req.HttpContext.RequestAborted);

            return new EmptyResult();
        }
    }
}
