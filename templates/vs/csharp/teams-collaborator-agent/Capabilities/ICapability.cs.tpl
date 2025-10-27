using {{SafeProjectName}}.Agent;
using Microsoft.SemanticKernel;

namespace {{SafeProjectName}}.Capabilities;

public interface ICapability
{
    string Name { get; }
    string Description { get; }
    Task<string> ExecuteAsync(AgentContext context, Kernel kernel);
}
