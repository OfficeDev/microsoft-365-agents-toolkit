import "@typespec/http";
import "@typespec/openapi3";
import "@microsoft/typespec-copilot-skills";

using TypeSpec.Http;
using TypespecCopilotSkills;
using TypespecCopilotSkills.Agents;

@agent(
  "{{appName}}",
  "Declarative agent created with Microsoft 365 Agents Toolkit"
)

@instructions("""
  You are a declarative agent and were created with Microsoft 365 Agents Toolkit.
""")

// Uncomment this part to add a conversation starter to the agent.
// This will be shown to the user when the agent is first created.
// @conversationStarter(#{
//   text: "<prompt tect>",
//   title: "<title>",
// })

namespace {{appName}} {
  // Uncomment this part to add an action in the agent.
  // @service
  // @skill({
  //   nameForHuman: "<API name>",
  //   descriptionForModel: "<API description>",
  // })

  // @server("<API URL endpoint>", "<API description>")
  // namespace <APIEndpointName> {
  //   /** action description **/
  //   @route("<path>")
  //   @get
  //   op <functionName>(
  //     /** param1 description **/
  //     @query param1: <type>
  //     /** param2 description **/
  //     @query param2: <type>,
  //   ): <returnType or ModelName>

  //   model <modelName> {
  //     <paramName1>: <type or Model>;
  //     <paramName2>: <type or Model>;
  //   }
  // }
}