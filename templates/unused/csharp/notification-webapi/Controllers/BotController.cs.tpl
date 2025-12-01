using Microsoft.Agents.Builder;
using Microsoft.Agents.Hosting.AspNetCore;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using {{SafeProjectName}}.Notification;

namespace {{SafeProjectName}}.Controllers
{
    [Route("api/messages")]
    [ApiController]
    [Authorize]
    public class BotController(IAgentHttpAdapter adapter, IAgent bot, NotificationBot notification) : ControllerBase
    {
        [HttpPost]
        public Task PostAsync(CancellationToken cancellationToken)
            => adapter.ProcessAsync(Request, Response, bot, cancellationToken);
    }
}