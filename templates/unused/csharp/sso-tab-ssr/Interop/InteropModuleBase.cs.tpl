using Microsoft.JSInterop;

namespace {{SafeProjectName}}.Interop;

public abstract class InteropModuleBase
{
    private readonly IJSRuntime _jsRuntime;
    private IJSObjectReference _module;

    protected IJSObjectReference Interop => _module;
    protected virtual string ModulePath { get; set; }

    public InteropModuleBase(IJSRuntime jsRuntime)
    {
        _jsRuntime = jsRuntime;
    }

    protected async Task InvokeVoidAsync(string functionName, params object[] args)
    {
        await ImportModuleAsync();

        await Interop.InvokeVoidAsync(functionName, args).AsTask();
    }

    protected async Task<T> InvokeAsync<T>(string functionName, params object[] args)
    {
        await ImportModuleAsync();

        return await Interop.InvokeAsync<T>(functionName, args).AsTask();
    }

    protected virtual async Task<IJSObjectReference> ImportModuleAsync()
    {
        if (_module == null)
        {
            _ = await ImportPrerequisiteModuleAsync("https://res.cdn.office.net/teams-js/2.22.0/js/MicrosoftTeams.min.js");
            _module = await _jsRuntime.InvokeAsync<IJSObjectReference>("import", ModulePath).AsTask();
        }

        return _module;
    }

    private Task<IJSObjectReference> ImportPrerequisiteModuleAsync(string url)
    {
        return _jsRuntime.InvokeAsync<IJSObjectReference>("import", url).AsTask();
    }
}