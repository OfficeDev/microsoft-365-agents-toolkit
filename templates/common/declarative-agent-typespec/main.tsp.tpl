import "@typespec/http";
import "@typespec/openapi3";
import "@microsoft/typespec-m365-copilot";
import "./actions.tsp";

using TypeSpec.Http;
using TypeSpec.M365.Copilot.Agents;
using TypeSpec.M365.Copilot.Actions;

@agent(
  "{{appName}}",
  "Declarative agent created with Microsoft 365 Agents Toolkit and TypeSpec for Microsoft 365 Copilot."
)

@instructions("""
  You are a declarative agent and were created with Microsoft 365 Agents Toolkit and TypeSpec for Microsoft 365 Copilot.
""")

// Uncomment this part to add a conversation starter to the agent.
// This will be shown to the user when the agent is first created.
// @conversationStarter(#{
//   title: "Get latest issues",
//   text: "Get the latest issues from GitHub" 
// })

namespace {{appName}} {  
  // Uncomment this part to add actions to the agent.
  // @service
  // @server(global.GitHubAPI.SERVER_URL)
  // @actions(global.GitHubAPI.ACTIONS_METADATA)
  // namespace GitHubAPIActions {
  //   op searchIssues is global.GitHubAPI.searchIssues;
  // }
}