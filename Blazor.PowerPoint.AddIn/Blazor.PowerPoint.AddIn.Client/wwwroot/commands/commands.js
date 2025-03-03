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

// Associate the functions with their named counterparts in the manifest XML.
Office.actions.associate("action", insertTextInPowerPoint);