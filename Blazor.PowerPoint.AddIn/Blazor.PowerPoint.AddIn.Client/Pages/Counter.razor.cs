/* Copyright(c) Maarten van Stam All rights reserved. */

using Microsoft.AspNetCore.Components;

namespace Blazor.PowerPoint.AddIn.Client.Pages;

public partial class Counter : ComponentBase
{
    private int currentCount = 0;

    private void IncrementCount()
    {
        currentCount++;
    }
}
