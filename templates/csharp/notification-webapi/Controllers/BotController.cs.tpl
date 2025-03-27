using Microsoft.Agents.BotBuilder;
using Microsoft.Agents.Hosting.AspNetCore;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.TeamsFx.Conversation;

namespace {{SafeProjectName}}.Controllers
{
    [Route("api/messages")]
    [ApiController]
    [Authorize]
    public class BotController(IBotHttpAdapter adapter, IBot bot, ConversationBot conversation) : ControllerBase
    {
        [HttpPost]
        public Task PostAsync(CancellationToken cancellationToken)
            => adapter.ProcessAsync(Request, Response, bot, cancellationToken);
    }
}