using Microsoft.JSInterop;

namespace Blazor.PowerPoint.AddIn.Client.Services;

/// <summary>
/// Command handler service for Server-side ribbon commands.
/// Methods are exposed to JavaScript via DotNetObjectReference through ServerBridge.
/// Executes on the server via SignalR.
/// </summary>
public class ServerCommandHandler
{
    [JSInvokable]
    public Task<string> SayHelloCounter(string name)
    {
        Console.WriteLine($"Invoking SayHelloCounter {name}");
        return Task.FromResult($"Hello Counter, {name} from the InteractiveServer Counter Page!");
    }
}
