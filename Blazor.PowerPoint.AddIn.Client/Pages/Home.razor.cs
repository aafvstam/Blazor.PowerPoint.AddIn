/* Copyright(c) Maarten van Stam. All rights reserved. Licensed under the MIT License. */
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices.JavaScript;
using System.Runtime.Versioning;

using Blazor.PowerPoint.AddIn.Client.Model;

using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;

namespace Blazor.PowerPoint.AddIn.Client.Pages;

[SupportedOSPlatform("browser")]
public partial class Home : ComponentBase
{
    private HostInformation hostInformation = new HostInformation();

    // Static field to cache the render mode when the component is rendered
    private static string? _cachedRenderMode;

    /// <summary>
    /// Gets the current render mode name based on the component's RendererInfo.
    /// </summary>
    private string CurrentRenderMode => RendererInfo.IsInteractive
        ? $"Interactive{RendererInfo.Name}"
        : "Static";

    [Inject, AllowNull]
    private IJSRuntime JSRuntime { get; set; }

    private IJSObjectReference JSModule { get; set; } = default!;

    protected override async Task OnAfterRenderAsync(bool firstRender)
    {
        if (firstRender)
        {
            // Cache the render mode for static method access
            _cachedRenderMode = CurrentRenderMode;

            hostInformation = await JSRuntime.InvokeAsync<HostInformation>("Office.onReady");

            Debug.WriteLine("Hit OnAfterRenderAsync in Home.razor.cs!");
            Console.WriteLine("Hit OnAfterRenderAsync in Home.razor.cs in Console!");
            JSModule = await JSRuntime.InvokeAsync<IJSObjectReference>("import", "./Pages/Home.razor.js");

            if (hostInformation.IsInitialized)
            {
                StateHasChanged();
            }
        }
    }

    /// <summary>
    /// Function to create a new slide in the PowerPoint presentation.
    /// </summary>
    private async Task CreateSlide() =>
        await JSModule.InvokeVoidAsync("createSlide");

    // Static JSImport method to call into Home.razor.js (synchronous - returns string, not Promise)
    [JSImport("sayHelloFromJs", "Home")]
    internal static partial string SayHelloFromJsFunction(string name);

    /// <summary>
    /// Static JSInvokable method called from ribbon button via DotNet.invokeMethodAsync.
    /// Imports the Home.razor.js module and calls into it.
    /// </summary>
    /// <param name="name">The name to greet.</param>
    /// <returns>A greeting string with render mode information.</returns>
    [JSInvokable]
    public static async Task<string> SayHelloHome(string name)
    {
        Console.WriteLine($"Invoking static SayHelloHome {name}");

        // Import the Home module (collocated JS file)
        await JSHost.ImportAsync("Home", "../Pages/Home.razor.js");

        // Call the JS function and get the result (synchronous call)
        var result = SayHelloFromJsFunction(name);

        // Use cached render mode if available, otherwise fall back to OperatingSystem check
        var renderMode = _cachedRenderMode 
            ?? (OperatingSystem.IsBrowser() ? "InteractiveWebAssembly" : "Unknown");

        return $"{result} from the {renderMode} Home Page!";
    }
}