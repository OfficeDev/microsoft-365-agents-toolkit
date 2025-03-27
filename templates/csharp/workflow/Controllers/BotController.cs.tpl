using Microsoft.Agents.BotBuilder;
using Microsoft.Agents.Hosting.AspNetCore;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace {{SafeProjectName}}.Controllers
{
    [Route("api/messages")]
    [ApiController]
    [Authorize]
    public class BotController(IBotHttpAdapter adapter, IBot bot) : ControllerBase
    {
        [HttpPost]
        public Task PostAsync(CancellationToken cancellationToken)
            => adapter.ProcessAsync(Request, Response, bot, cancellationToken);
    }
}
