/* Copyright(c) Maarten van Stam All rights reserved. */

namespace Blazor.PowerPoint.AddIn.Components.Layout
{
    public partial class MainLayout
    {
        public string TrademarkMessage1 { get; set; } = "Copyright © " + @DateTime.Now.Year + " Maarten van Stam.";
        public string TrademarkMessage2 { get; set; } = "All rights reserved.";

        public string FrameworkDescription { get; set; } = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription;
    }
}