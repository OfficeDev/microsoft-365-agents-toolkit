import "@typespec/http"; 
import "@typespec/openapi3"; 
import "@microsoft/typespec-copilot-skills"; 

using TypeSpec.Http; 
using TypespecCopilotSkills; 
using TypespecCopilotSkills.Agents; 

@agent(
  "<name>",
  "<agent description>" 
) 

@instructions(""" 
  <agent instructions>
""") 

@conversationStarter(#{
  text: "<prompt tect>",
  title: "<title>",
}) 

namespace <AgentName> {
  @service
  @skill({ 
    nameForHuman: "<API name>", 
    descriptionForModel: "<API description>", 
  })

  @server("<API URL endpoint>", "<API description>")
  namespace <APIEndpointName> {
    @route("<path>")
    @get
    op <functionName>(
      /** param1 description **/
      @query param1: <type>
      /** param2 description **/
      @query param2: <type>,
    ): <returnType or ModelName>

    model <modelName> {
      <paramName1>: <type or Model>;
      <paramName2>: <type or Model>;
    }
  }
}