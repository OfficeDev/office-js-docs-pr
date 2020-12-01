---
title: Insert and delete slides in a PowerPoint presentation
description: 'Learn how to insert slides from one presentation into another and how to delete slides.'
ms.date: 12/04/2020
localization_priority: Normal
---

# Insert and delete slides in a PowerPoint presentation

A PowerPoint add-in can insert slides from one presentation (in base64 format) into the current presentation by using PowerPoint's application-specific JavaScript library. You can control whether the inserted slides keep the formatting of the source presentation or the formatting of the target presentation. You can also delete slides from the presentation.

There are two major steps to inserting slides from one presentation into another.

1. Convert the source presentation file (.pptx) into a base64-formatted file.
1. Use the `insertSlidesFromBase64` method to insert one or more slides from the base64 file into the current presentation.

## Convert the source presentation to base64

There are many ways to convert a file to base64. Which programming language and library you use, and whether to convert on the server-side of your add-in or the client-side is determined by your scenario. Most commonly, you'll do the conversion in JavaScript on the client-side by using a [FileReader](https://developer.mozilla.org/en-US/docs/Web/API/FileReader) object. The following is an example.

1. Begin by getting a reference to the source PowerPoint file. In this example, we will use an `<input>` control of type `file` to prompt the user to choose a file. Add the following markup to the add-in page.

    ```html
    <section>
        <p>Select a PowerPoint presentation from which to insert slides</p>
        <form>
            <input type="file" id="file" />
        </form>
    </section>
    ```

    This markup adds the UI in the following screenshot to the page:

    ![Screenshot showing an HTML file type input control preceded by an instructional sentence reading "Select a PowerPoint presentation from which to insert slides". The control consists of a button labelled "Choose file" followed by the sentence "No file chosen".](../images/powerpoint-html-file-input-control.png)

2. Add the following code to the add-in's JavaScript to assign a function to the input control's `change` event. (You create the `storeFileAsBase64` in the next step.)

    ```javascript
    $("#file").change(storeFileAsBase64);
    ```

3. Add the following code. About this code, note the following:

    - The `reader.readAsDataURL` method converts the file to base64 and stores it in the `reader.result` property. When the method completes, it triggers the `onload` event handler.
    - The `onload` event handler trims metadata off of the encoded file and stores the encoded string in a global variable.
    - The base64-encoded string is stored globally because it will be read by another function that you create in a later step.

    ```javascript
    let chosenFileBase64;

    async function storeFileAsBase64() {
        const reader = new FileReader();

        reader.onload = async (event) => {
            const startIndex = reader.result.toString().indexOf("base64,");
            const copyBase64 = reader.result.toString().substr(startIndex + 7);

            chosenFileBase64 = copyBase64;
        };

        const myFile = document.getElementById("file") as HTMLInputElement;
        reader.readAsDataURL(myFile.files[0]);
    }
    ```

## Insert slides with insertSlidesFromBase64

You insert slides from another PowerPoint presentation into the current presentation with the [Presentation.insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-) method. 




## Delete slides

