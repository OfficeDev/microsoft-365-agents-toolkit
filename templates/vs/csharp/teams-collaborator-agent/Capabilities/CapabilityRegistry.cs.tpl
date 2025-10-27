using Microsoft.Extensions.Logging;
using Microsoft.SemanticKernel;

namespace {{SafeProjectName}}.Capabilities;

public interface ICapabilityRegistry
{
    void RegisterCapability(ICapability capability);
    ICapability? GetCapability(string name);
    IEnumerable<ICapability> GetAllCapabilities();
}

public class CapabilityRegistry : ICapabilityRegistry
{
    private readonly Dictionary<string, ICapability> _capabilities = new();
    private readonly Kernel _kernel;
    private readonly ILogger<CapabilityRegistry> _logger;

    public CapabilityRegistry(Kernel kernel, ILogger<CapabilityRegistry> logger)
    {
        _kernel = kernel;
        _logger = logger;
    }

    public void RegisterCapability(ICapability capability)
    {
        _capabilities[capability.Name] = capability;
        _logger.LogInformation($"✅ Registered capability: {capability.Name}");
    }

    public ICapability? GetCapability(string name)
    {
        _capabilities.TryGetValue(name, out var capability);
        return capability;
    }

    public IEnumerable<ICapability> GetAllCapabilities()
    {
        return _capabilities.Values;
    }
}
