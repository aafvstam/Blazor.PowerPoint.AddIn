/* Copyright(c) Maarten van Stam. All rights reserved. Licensed under the MIT License. */
using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;

using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;

namespace Blazor.PowerPoint.AddIn.Client.Pages;

public partial class Home : ComponentBase
{
    [Inject, AllowNull]
    private IJSRuntime JSRuntime { get; set; }
    private IJSObjectReference JSModule { get; set; } = default!;

    protected override async Task OnAfterRenderAsync(bool firstRender)
    {
        if (firstRender)
        {
            Debug.WriteLine("Hit OnAfterRenderAsync in Home.razor.cs!");
            Console.WriteLine("Hit OnAfterRenderAsync in Home.razor.cs in Console!");
            JSModule = await JSRuntime.InvokeAsync<IJSObjectReference>("import", "./Pages/Home.razor.js");
        }
    }

    /// <summary>
    /// Basic function to invoke inserting `Hello world!` text.
    /// </summary>
    private async Task HelloButton() =>
        await JSModule.InvokeVoidAsync("helloButton");

    [JSInvokable]
    public static string SayHelloHome(string name)
    {
        return $"Hello Index, {name} from Home!";
    }
}