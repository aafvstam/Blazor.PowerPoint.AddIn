/* Copyright(c) Maarten van Stam. All rights reserved. Licensed under the MIT License. */
using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;

namespace Blazor.PowerPoint.AddIn.Client.Pages;

public partial class Counter : ComponentBase
{
    private int currentCount = 0;

    private void IncrementCount()
    {
        currentCount++;
    }

    [JSInvokable]
    public static Task<string> SayHelloCounter(string name)
    {
        return Task.FromResult($"Hello Home, {name} from Counter Page!");
    }
}
