/* Copyright(c) Maarten van Stam. All rights reserved. Licensed under the MIT License. */

/**
 * Basic function to show how to insert a value into a PowerPoint Presentation.
 */
console.log("Loading Home.razor.js");

export async function helloButton() {

    console.log("We are now entering function: helloButton");

     /**
     * Insert your PowerPoint code here
     */
    const options: Office.SetSelectedDataOptions = { coercionType: Office.CoercionType.Text };

    await Office.context.document.setSelectedDataAsync(" ", options);
    await Office.context.document.setSelectedDataAsync("Hello World!", options);
}