/*
 * Copyright (c) Maarten van Stam. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 * 
 */

console.log("Loading command.js");

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
async function insertTextInPowerPoint(event) {

    console.log("In insertTextInPowerPoint");

    try {
        await PowerPoint.run(async (context) => {
            const slide = context.presentation.getSelectedSlides().getItemAt(0);
            const textBox = slide.shapes.addTextBox("Hello World");
            textBox.fill.setSolidColor("white");
            textBox.lineFormat.color = "black";
            textBox.lineFormat.weight = 1;
            textBox.lineFormat.dashStyle = PowerPoint.ShapeLineDashStyle.solid;
            await context.sync();
        });
    } catch (error) {
        console.log();
        console.log("Error call : " + error.message);
    }
    finally {
        console.log("Finish insertTextInPowerPoint");
    }

    // Be sure to indicate when the add-in command function is complete
    event.completed();
}

/**
 * Writes the text from the Home Blazor Page to the PowerPoint slide
 * @param {any} event
 */
async function callBlazorOnHome(event) {
    console.log("Running callBlazorOnHome");


    try {
        let name = "Initializing";

        try {

            //--------------------------------------------------------------------------------
            // TODO: Dotnet seems to be undefined on first run ... need to investigate
            //--------------------------------------------------------------------------------
            try {

                // Call JSInvokable Function here ...
                name = await DotNet.invokeMethodAsync("Blazor.PowerPoint.AddIn.Client", "SayHello", "Blazor Fan");

            } catch (err) {
                name = err.message;
                console.error("Error during DotNet invocation: " + err.message);
            }

            console.log("Finished Initializing: " + name)
        }
        catch (err) {
            name = err.message;
        }

        await PowerPoint.run(async (context) => {
            const slide = context.presentation.getSelectedSlides().getItemAt(0);
            const textBox = slide.shapes.addTextBox(name);
            textBox.fill.setSolidColor("white");
            textBox.lineFormat.color = "black";
            textBox.lineFormat.weight = 1;
            textBox.lineFormat.dashStyle = PowerPoint.ShapeLineDashStyle.solid;
            await context.sync();
        });
    }
    catch (err) {
        console.log();
        console.log("Error call : " + err.message);
    }
    finally {
        console.log("Finish callBlazorOnHome");
    }

    // Calling event.completed is required. event.completed lets the platform know that processing has completed.
    event.completed();
}

// Associate the functions with their named counterparts in the manifest XML.
Office.actions.associate("insertTextInPowerPoint", insertTextInPowerPoint);
Office.actions.associate("callBlazorOnHome", callBlazorOnHome);