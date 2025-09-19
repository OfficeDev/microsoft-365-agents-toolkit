using Microsoft.Teams.AI.State;
using Microsoft.Teams.AI;
using {{SafeProjectName}}.Models;

namespace {{SafeProjectName}}
{
    public class APIBot : Application<AppState>
    {
        public APIBot(ApplicationOptions<AppState> options) : base(options)
        {
            // Registering action handlers that will be hooked up to the planner.
            AI.ImportActions(new APIActions());
        }
    }
}