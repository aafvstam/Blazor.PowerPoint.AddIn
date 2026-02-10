"use strict";
/*
 * Copyright (c) Maarten van Stam. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
console.log("Loading command.js");
/**
 * Inserts "Hello World" box in the PowerPoint presentation.
 * This function demonstrates basic Office JavaScript API usage without Blazor interop.
 *
 * @param event - The Office add-in command event object
 * @returns A promise that resolves when the text insertion is complete
 */
async function insertTextInPowerPoint(event) {
    console.log("In insertTextInPowerPoint");
    try {
        await PowerPoint.run(async (context) => {
            const slide = context.presentation.getSelectedSlides().getItemAt(0);
            const textBox = slide.shapes.addTextBox("Hello World", {
                left: 255,
                top: 25,
                height: 50,
                width: 250
            });
            textBox.fill.setSolidColor("white");
            textBox.lineFormat.color = "black";
            textBox.lineFormat.weight = 1;
            textBox.lineFormat.dashStyle = PowerPoint.ShapeLineDashStyle.solid;
            // Align text in the middle of the text box
            textBox.textFrame.textRange.paragraphFormat.horizontalAlignment = PowerPoint.ParagraphHorizontalAlignment.center;
            await context.sync();
        });
    }
    catch (error) {
        const errorMessage = error instanceof Error ? error.message : String(error);
        console.error("Error in insertTextInPowerPoint:", errorMessage);
    }
    finally {
        console.log("Finish insertTextInPowerPoint");
    }
    // Be sure to indicate when the add-in command function is complete
    if (event && typeof event.completed === 'function') {
        event.completed();
    }
}
/**
 * Writes the text from the Home Blazor Page to the PowerPoint slide
 * @param {any} event
 */
async function callBlazorOnHome(event) {
    // Implement your custom code here. The following code is a simple PowerPoint example.  
    try {
        console.log("Running callBlazorOnHome");
        await callStaticLocalComponentMethodinit("SayHelloHome");
    }
    catch (error) {
        const errorMessage = error instanceof Error ? error.message : String(error);
        console.error("Error in callBlazorOnHome:", errorMessage);
    }
    finally {
        console.log("Finish callBlazorOnHome");
    }
    // Be sure to indicate when the add-in command function is complete
    if (event && typeof event.completed === 'function') {
        event.completed();
    }
}
/**
 * Writes the text from the Counter Blazor Page to the PowerPoint slide
 * @param {any} event
 */
async function callBlazorOnCounter(event) {
    try {
        console.log("Running callBlazorOnCounter");
        await callStaticLocalComponentMethodinit("SayHelloCounter");
    }
    catch (error) {
        const errorMessage = error instanceof Error ? error.message : String(error);
        console.error("Error in callBlazorOnCounter:", errorMessage);
    }
    finally {
        console.log("Finish callBlazorOnCounter");
    }
    // Calling event.completed is required. event.completed lets the platform know that processing has completed.
    if (event && typeof event.completed === 'function') {
        event.completed();
    }
}
/**
 * Checks if the .NET runtime is loaded and invokes a .NET method to retrieve a string.
 * The string is then inserted into a PowerPoint slide as a text box.
 * and some format is added to the text box.
 *
 * @param {string} methodname - The name of the .NET method to invoke.
 */
async function callStaticLocalComponentMethodinit(methodname) {
    console.log("In callStaticLocalComponentMethodinit");
    try {
        let name = "Initializing";
        try {
            const dotnetloaded = await preloadDotNet();
            name = "something";
            if (dotnetloaded === true) {
                name = "Dotnet Loaded";
                // Call JSInvokable Function here ...
                name = await DotNet.invokeMethodAsync("Blazor.PowerPoint.AddIn.Client", methodname, "Blazor Fan");
            }
            else {
                name = "Init DotNet Failed, methodname: " + methodname;
            }
        }
        catch (error) {
            const errorMessage = error instanceof Error ? error.message : String(error);
            name = errorMessage;
            console.error("Error during DotNet invocation: " + name);
        }
        await PowerPoint.run(async (context) => {
            const slide = context.presentation.getSelectedSlides().getItemAt(0);
            const textBox = slide.shapes.addTextBox(name, {
                left: 255,
                top: 50,
                height: 50,
                width: 450
            });
            textBox.fill.setSolidColor("white");
            textBox.lineFormat.color = "black";
            textBox.lineFormat.weight = 1;
            textBox.lineFormat.dashStyle = PowerPoint.ShapeLineDashStyle.solid;
            // Align text in the middle of the text box
            textBox.textFrame.textRange.paragraphFormat.horizontalAlignment = PowerPoint.ParagraphHorizontalAlignment.center;
            await context.sync();
        });
        console.log("Finished Initializing: " + name);
    }
    catch (error) {
        const errorMessage = error instanceof Error ? error.message : String(error);
        console.error("Error in callStaticLocalComponentMethodinit:", errorMessage);
    }
    finally {
        console.log("Finish callStaticLocalComponentMethodinit");
    }
}
/**
 * Waits for the .NET runtime to be ready.
 *
 * Uses a promise-based approach where the Blazor module signals readiness
 * via afterWebAssemblyStarted, eliminating the need for polling.
 *
 * @param timeoutMs - Maximum time to wait for .NET to be ready (default: 10000ms)
 * @returns {Promise<boolean>} Returns true if the .NET runtime is ready, false if timeout.
 */
async function preloadDotNet(timeoutMs = 10000) {
    console.log("In preloadDotNet");
    try {
        const dotNetReadyPromise = window.dotNetReady;
        if (!dotNetReadyPromise) {
            console.error("dotNetReady promise not found - Blazor module may not be loaded");
            return false;
        }
        // Race between the ready promise and a timeout
        const timeoutPromise = new Promise((_, reject) => setTimeout(() => reject(new Error("Timeout waiting for .NET runtime")), timeoutMs));
        await Promise.race([dotNetReadyPromise, timeoutPromise]);
        console.log(".NET runtime is ready");
        return true;
    }
    catch (error) {
        const errorMessage = error instanceof Error ? error.message : String(error);
        console.error("Error in preloadDotNet: " + errorMessage);
        return false;
    }
    finally {
        console.log("Finish preloadDotNet");
    }
}
// Associate the functions with their named counterparts in the manifest XML.
Office.actions.associate("insertTextInPowerPoint", insertTextInPowerPoint);
Office.actions.associate("callBlazorOnHome", callBlazorOnHome);
Office.actions.associate("callBlazorOnCounter", callBlazorOnCounter);
//# sourceMappingURL=commands.js.map