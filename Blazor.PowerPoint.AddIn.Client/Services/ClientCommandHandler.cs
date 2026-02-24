using Microsoft.JSInterop;

namespace Blazor.PowerPoint.AddIn.Client.Services;

/// <summary>
/// Command handler service for WebAssembly-side ribbon commands.
/// Methods are exposed to JavaScript via DotNetObjectReference through WasmBridge.
/// Executes in the browser's WebAssembly runtime.
/// </summary>
public class ClientCommandHandler
{
    [JSInvokable]
    public Task<string> SayHelloHome(string name)
    {
        return Task.FromResult($"Hello Home, {name} from the InteractiveWebAssembly Home Page!");
    }
}
