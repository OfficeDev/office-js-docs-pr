---
title: Word add-in tutorial
description: In this tutorial, you'll build a Word add-in that inserts (and replaces) text ranges, paragraphs, images, HTML, tables, and content controls. You'll also learn how to format text and how to insert (and replace) content in content controls.
ms.date: 12/11/2025
ms.service: word
#Customer intent: As a developer, I want to build a Word add-in that can interact with content in a Word document.
ms.localizationpriority: high
---

# Tutorial: Create a Word task pane add-in

In this tutorial, you'll create a Word task pane add-in that:

> [!div class="checklist"]
>
> - Inserts a range of text
> - Formats text
> - Replaces text and inserts text in various locations
> - Inserts images, HTML, and tables
> - Creates and updates content controls

> [!TIP]
> If you've already completed the [Build your first Word task pane add-in](../quickstarts/word-quickstart-yo.md) quick start, and want to use that project as a starting point for this tutorial, go directly to the [Insert a range of text](#insert-a-range-of-text) section to start this tutorial.
>
> If you want a completed version of this tutorial, visit the [Office Add-ins samples repo on GitHub](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/tutorials/word-tutorial).

## Prerequisites

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## Create your add-in project

Decide the type of manifest that you'd like to use, either the **unified manifest for Microsoft 365** or the **add-in only manifest**. To learn more about them, see [Office Add-ins manifest](../develop/add-in-manifests.md).

# [Unified manifest for Microsoft 365 (preview)](#tab/jsonmanifest)

> [!NOTE]
> Using the unified manifest for Microsoft 365 with Word add-ins is in public developer preview. The unified manifest for Microsoft 365 shouldn't be used in production Word add-ins. We invite you to try it out in test or development environments. For more information, see the [Microsoft 365 app manifest schema reference](/microsoft-365/extensibility/schema).

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- **Choose a project type:** `Excel, PowerPoint, and/or Word Task Pane with unified manifest for Microsoft 365 (preview)`
- **What do you want to name your add-in?** `My Office Add-in`
- **Which Office client application would you like to support?** `Word`

:::image type="content" source="../images/yo-office-powerpoint-json-manifest-preview.png" alt-text="The prompts and answers for the Yeoman generator in a command line interface when the unified manifest is selected.":::

# [Add-in only manifest](#tab/xmlmanifest)

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- **Choose a project type:** `Office Add-in Task Pane project`
- **Choose a script type:** `TypeScript`
- **What do you want to name your add-in?** `My Office Add-in`
- **Which Office client application would you like to support?** `Word`

:::image type="content" source="../images/yo-office-word-xml-manifest-ts.png" alt-text="The prompts and answers for the Yeoman generator in a command line interface when the add-in only manifest is selected.":::

---

After you complete the wizard, the generator creates the project and installs supporting Node components.

## Insert a range of text

In this step of the tutorial, you'll programmatically test that your add-in supports the user's current version of Word, and then insert a paragraph into the document.

### Code the add-in

1. Open the project in your code editor.

1. Open the file **./src/taskpane/taskpane.html**. This file contains the HTML markup for the task pane.

1. Locate the `<main>` element and delete all lines that appear after the opening `<main>` tag and before the closing `</main>` tag.

1. Add the following markup immediately after the opening `<main>` tag.

    ```html
    <button class="ms-Button" id="insert-paragraph">Insert Paragraph</button><br/><br/>
    ```

1. Open the file **./src/taskpane/taskpane.ts**. This file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office client application. Replace the entire contents with the following code and save the file.

    ```js
    /*
     * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
     * See LICENSE in the project root for license information.
     */
    
    /* global document, Office, Word */
    
    Office.onReady((info) => {
      if (info.host === Office.HostType.Word) {
        document.getElementById("sideload-msg").style.display = "none";
        document.getElementById("app-body").style.display = "flex";
      }
    });
    ```

1. Within the `Office.onReady` function call, locate the line `if (info.host === Office.HostType.Word) {` and add the following code immediately after that line. Note:

    - This code adds an event handler for the `insert-paragraph` button.
    - The `insertParagraph` function is wrapped in a call to `tryCatch` (both functions will be added in the next step). This allows any errors generated by the Office JavaScript API layer to be handled separately from your service code.

    ```js
    // Assign event handlers and other initialization logic.
    document.getElementById("insert-paragraph").onclick = () => tryCatch(insertParagraph);
    ```

1. Add the following functions to the end of the file. Note:

   - Your Word.js business logic will be added to the function passed to `Word.run`. This logic doesn't execute immediately. Instead, it's added to a queue of pending commands.

   - The `context.sync` method sends all queued commands to Word for execution.

   - The `tryCatch` function will be used by all the functions interacting with the workbook from the task pane. Catching Office JavaScript errors in this fashion is a convenient way to generically handle uncaught errors.

    ```js
    async function insertParagraph() {
        await Word.run(async (context) => {

            // TODO1: Queue commands to insert a paragraph into the document.

            await context.sync();
        });
    }

    /** Default helper for invoking an action and handling errors. */
    async function tryCatch(callback) {
        try {
            await callback();
        } catch (error) {
            // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
            console.error(error);
        }
    }
    ```

1. Within the `insertParagraph()` function, replace `TODO1` with the following code. Note:

   - The first parameter to the `insertParagraph` method is the text for the new paragraph.

   - The second parameter is the location within the body where the paragraph will be inserted. Other options for insert paragraph, when the parent object is the body, are "End" and "Replace".

    ```js
    const docBody = context.document.body;
    docBody.insertParagraph("Office has several versions, including Office 2021, Microsoft 365 subscription, and Office on the web.",
                            Word.InsertLocation.start);
    ```

1. Save all your changes to the project.

### Test the add-in

1. Complete the following steps to start the local web server and sideload your add-in.

    [!INCLUDE [alert use https](../includes/alert-use-https.md)]

    > [!TIP]
    > If you're testing your add-in on Mac, run the following command in the root directory of your project before proceeding. When you run this command, the local web server starts.
    >
    > ```command&nbsp;line
    > npm run dev-server
    > ```

    - To test your add-in in Word, run the following command in the root directory of your project. This starts the local web server (if it isn't already running) and opens Word with your add-in loaded.

        ```command&nbsp;line
        npm start
        ```

    - To test your add-in in Word on the web, run the following command in the root directory of your project. When you run this command, the local web server starts. Replace "{url}" with the URL of a Word document on your OneDrive or a SharePoint library to which you have permissions.

        [!INCLUDE [npm start on web command syntax](../includes/start-web-sideload-instructions.md)]

1. In Word, if the "My Office Add-in" task pane isn't already open, choose the **Home** tab, and then choose the **Show Task Pane** button on the ribbon to open the add-in task pane.

    :::image type="content" source="../images/word-quickstart-add-in-2b.png" alt-text="The Show Task Pane button highlighted in Word.":::

1. In the task pane, choose the **Insert Paragraph** button.

1. Make a change in the paragraph.

1. Choose the **Insert Paragraph** button again. Note that the new paragraph appears above the previous one because the `insertParagraph` method is inserting at the start of the document's body.

    :::image type="content" source="../images/word-tutorial-insert-paragraph-2.png" alt-text="The Insert Paragraph button in the add-in.":::

1. [!include[Instructions to stop web server and uninstall dev add-in](../includes/stop-uninstall-dev-add-in.md)]

## Format text

In this step of the tutorial, you'll apply a built-in style to text, apply a custom style to text, and change the font of text.

### Apply a built-in style to text

1. Open the file **./src/taskpane/taskpane.html**.

1. Locate the `<button>` element for the `insert-paragraph` button, and add the following markup after that line.

    ```html
    <button class="ms-Button" id="apply-style">Apply Style</button><br/><br/>
    ```

1. Open the file **./src/taskpane/taskpane.ts**.

1. Within the `Office.onReady` function call, locate the line that assigns a click handler to the `insert-paragraph` button, and add the following code after that line.

    ```js
    document.getElementById("apply-style").onclick = () => tryCatch(applyStyle);
    ```

1. Add the following function to the end of the file.

    ```js
    async function applyStyle() {
        await Word.run(async (context) => {

            // TODO1: Queue commands to style text.

            await context.sync();
        });
    }
    ```

1. Within the `applyStyle()` function, replace `TODO1` with the following code. Note that the code applies a style to a paragraph, but styles can also be applied to ranges of text.

    ```js
    const firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.styleBuiltIn = Word.Style.intenseReference;
    ```

### Apply a custom style to text

1. Open the file **./src/taskpane/taskpane.html**.

1. Locate the `<button>` element for the `apply-style` button, and add the following markup after that line.

    ```html
    <button class="ms-Button" id="apply-custom-style">Apply Custom Style</button><br/><br/>
    ```

1. Open the file **./src/taskpane/taskpane.ts**.

1. Within the `Office.onReady` function call, locate the line that assigns a click handler to the `apply-style` button, and add the following code after that line.

    ```js
    document.getElementById("apply-custom-style").onclick = () => tryCatch(applyCustomStyle);
    ```

1. Add the following function to the end of the file.

    ```js
    async function applyCustomStyle() {
        await Word.run(async (context) => {

            // TODO1: Queue commands to apply the custom style.

            await context.sync();
        });
    }
    ```

1. Within the `applyCustomStyle()` function, replace `TODO1` with the following code. Note that the code applies a custom style that does not exist yet. You'll create a style with the name **MyCustomStyle** in the [Test the add-in](#test-the-add-in-1) step.

    ```js
    const lastParagraph = context.document.body.paragraphs.getLast();
    lastParagraph.style = "MyCustomStyle";
    ```

1. Save all your changes to the project.

### Change the font of text

1. Open the file **./src/taskpane/taskpane.html**.

1. Locate the `<button>` element for the `apply-custom-style` button, and add the following markup after that line.

    ```html
    <button class="ms-Button" id="change-font">Change Font</button><br/><br/>
    ```

1. Open the file **./src/taskpane/taskpane.ts**.

1. Within the `Office.onReady` function call, locate the line that assigns a click handler to the `apply-custom-style` button, and add the following code after that line.

    ```js
    document.getElementById("change-font").onclick = () => tryCatch(changeFont);
    ```

1. Add the following function to the end of the file.

    ```js
    async function changeFont() {
        await Word.run(async (context) => {

            // TODO1: Queue commands to apply a different font.

            await context.sync();
        });
    }
    ```

1. Within the `changeFont()` function, replace `TODO1` with the following code. Note that the code gets a reference to the second paragraph by using the `ParagraphCollection.getFirst` method chained to the `Paragraph.getNext` method.

    ```js
    const secondParagraph = context.document.body.paragraphs.getFirst().getNext();
    secondParagraph.font.set({
            name: "Courier New",
            bold: true,
            size: 18
        });
    ```

1. Save all your changes to the project.

### Test the add-in

1. [!include[Start server and sideload add-in instructions](../includes/tutorial-word-start-server.md)]

1. If the add-in task pane isn't already open in Word, go to the **Home** tab and choose the **Show Task Pane** button on the ribbon to open it.

1. Be sure there are at least three paragraphs in the document. You can choose the **Insert Paragraph** button three times. *Check carefully that there's no blank paragraph at the end of the document. If there is, delete it.*

1. In Word, create a [custom style](https://support.microsoft.com/office/d38d6e47-f6fc-48eb-a607-1eb120dec563) named "MyCustomStyle". It can have any formatting that you want.

1. Choose the **Apply Style** button. The first paragraph will be styled with the built-in style **Intense Reference**.

1. Choose the **Apply Custom Style** button. The last paragraph will be styled with your custom style. (If nothing seems to happen, the last paragraph might be blank. If so, add some text to it.)

1. Choose the **Change Font** button. The font of the second paragraph changes to 18 pt., bold, Courier New.

    :::image type="content" source="../images/word-tutorial-apply-styles-and-font-2.png" alt-text="The results of applying the styles and fonts defined for the add-in buttons Apply Style, Apply Custom Style, and Change font.":::

## Replace text and insert text

In this step of the tutorial, you'll add text inside and outside of selected ranges of text, and replace the text of a selected range.

### Add text inside a range

1. Open the file **./src/taskpane/taskpane.html**.

1. Locate the `<button>` element for the `change-font` button, and add the following markup after that line.

    ```html
    <button class="ms-Button" id="insert-text-into-range">Insert Abbreviation</button><br/><br/>
    ```

1. Open the file **./src/taskpane/taskpane.ts**.

1. Within the `Office.onReady` function call, locate the line that assigns a click handler to the `change-font` button, and add the following code after that line.

    ```js
    document.getElementById("insert-text-into-range").onclick = () => tryCatch(insertTextIntoRange);
    ```

1. Add the following function to the end of the file.

    ```js
    async function insertTextIntoRange() {
        await Word.run(async (context) => {

            // TODO1: Queue commands to insert text into a selected range.

            // TODO2: Load the text of the range and sync so that the
            //        current range text can be read.

            // TODO3: Queue commands to repeat the text of the original
            //        range at the end of the document.

            await context.sync();
        });
    }
    ```

1. Within the `insertTextIntoRange()` function, replace `TODO1` with the following code. Note:

   - The function is intended to insert the abbreviation ["(M365)"] into the end of the Range whose text is "Microsoft 365". It makes a simplifying assumption that the string is present and the user has selected it.

   - The first parameter of the `Range.insertText` method is the string to insert into the `Range` object.

   - The second parameter specifies where in the range the additional text should be inserted. Besides "End", the other possible options are "Start", "Before", "After", and "Replace".

   - The difference between "End" and "After" is that "End" inserts the new text inside the end of the existing range, but "After" creates a new range with the string and inserts the new range after the existing range. Similarly, "Start" inserts text inside the beginning of the existing range and "Before" inserts a new range. "Replace" replaces the text of the existing range with the string in the first parameter.

   - You saw in an earlier stage of the tutorial that the `insert*` methods of the body object don't have the "Before" and "After" options. This is because you can't put content outside of the document's body.

    ```js
    const doc = context.document;
    const originalRange = doc.getSelection();
    originalRange.insertText(" (M365)", Word.InsertLocation.end);
    ```

1. We'll skip over `TODO2` until the next section. Within the `insertTextIntoRange()` function, replace `TODO3` with the following code. This code is similar to the code you created in the first stage of the tutorial, except that now you are inserting a new paragraph at the end of the document instead of at the start. This new paragraph will demonstrate that the new text is now part of the original range.

    ```js
    doc.body.insertParagraph("Original range: " + originalRange.text, Word.InsertLocation.end);
    ```

### Add code to fetch document properties into the task pane's script objects

In all previous functions in this tutorial, you queued commands to *write* to the Office document. Each function ended with a call to the `context.sync()` method which sends the queued commands to the document to be executed. But the code you added in the last step calls the `originalRange.text` property, and this is a significant difference from the earlier functions you wrote, because the `originalRange` object is only a proxy object that exists in your task pane's script. It doesn't know what the actual text of the range in the document is, so its `text` property can't have a real value. It's necessary to first fetch the text value of the range from the document and use it to set the value of `originalRange.text`. Only then can `originalRange.text` be called without causing an exception to be thrown. This fetching process has three steps.

1. Queue a command to load (that is, fetch) the properties that your code needs to read.

1. Call the context object's `sync` method to send the queued command to the document for execution and return the requested information.

1. Because the `sync` method is asynchronous, ensure that it has completed before your code calls the properties that were fetched.

The following step must be completed whenever your code needs to *read* information from the Office document.

1. Within the `insertTextIntoRange()` function, replace `TODO2` with the following code.
  
    ```js
    originalRange.load("text");
    await context.sync();
    ```

When you're done, the entire function should look like the following:

```js
async function insertTextIntoRange() {
    await Word.run(async (context) => {

        const doc = context.document;
        const originalRange = doc.getSelection();
        originalRange.insertText(" (M365)", Word.InsertLocation.end);

        originalRange.load("text");
        await context.sync();

        doc.body.insertParagraph("Original range: " + originalRange.text, Word.InsertLocation.end);

        await context.sync();
    });
}
```

### Add text between ranges

1. Open the file **./src/taskpane/taskpane.html**.

1. Locate the `<button>` element for the `insert-text-into-range` button, and add the following markup after that line.

    ```html
    <button class="ms-Button" id="insert-text-outside-range">Add Version Info</button><br/><br/>
    ```

1. Open the file **./src/taskpane/taskpane.ts**.

1. Within the `Office.onReady` function call, locate the line that assigns a click handler to the `insert-text-into-range` button, and add the following code after that line.

    ```js
    document.getElementById("insert-text-outside-range").onclick = () => tryCatch(insertTextBeforeRange);
    ```

1. Add the following function to the end of the file.

    ```js
    async function insertTextBeforeRange() {
        await Word.run(async (context) => {

            // TODO1: Queue commands to insert a new range before the
            //        selected range.

            // TODO2: Load the text of the original range and sync so that the
            //        range text can be read and inserted.

        });
    }
    ```

1. Within the `insertTextBeforeRange()` function, replace `TODO1` with the following code. Note:

   - The function is intended to add a range whose text is "Office 2024, " before the range with text "Microsoft 365". It makes an assumption that the string is present and the user has selected it.

   - The first parameter of the `Range.insertText` method is the string to add.

   - The second parameter specifies where in the range the additional text should be inserted. For more details about the location options, see the previous discussion of the `insertTextIntoRange` function.

    ```js
    const doc = context.document;
    const originalRange = doc.getSelection();
    originalRange.insertText("Office 2024, ", Word.InsertLocation.before);
    ```

1. Within the `insertTextBeforeRange()` function, replace `TODO2` with the following code.

     ```js
    originalRange.load("text");
    await context.sync();

    // TODO3: Queue commands to insert the original range as a
    //        paragraph at the end of the document.

    // TODO4: Make a final call of context.sync here and ensure
    //        that it runs after the insertParagraph has been queued.
    ```

1. Replace `TODO3` with the following code. This new paragraph will demonstrate the fact that the new text is ***not*** part of the original selected range. The original range still has only the text it had when it was selected.

    ```js
    doc.body.insertParagraph("Current text of original range: " + originalRange.text, Word.InsertLocation.end);
    ```

1. Replace `TODO4` with the following code.

    ```js
    await context.sync();
    ```

### Replace the text of a range

1. Open the file **./src/taskpane/taskpane.html**.

1. Locate the `<button>` element for the `insert-text-outside-range` button, and add the following markup after that line.

    ```html
    <button class="ms-Button" id="replace-text">Change Quantity Term</button><br/><br/>
    ```

1. Open the file **./src/taskpane/taskpane.ts**.

1. Within the `Office.onReady` function call, locate the line that assigns a click handler to the `insert-text-outside-range` button, and add the following code after that line.

    ```js
    document.getElementById("replace-text").onclick = () => tryCatch(replaceText);
    ```

1. Add the following function to the end of the file.

    ```js
    async function replaceText() {
        await Word.run(async (context) => {

            // TODO1: Queue commands to replace the text.

            await context.sync();
        });
    }
    ```

1. Within the `replaceText()` function, replace `TODO1` with the following code. Note that the function is intended to replace the string "several" with the string "many". It makes a simplifying assumption that the string is present and the user has selected it.

    ```js
    const doc = context.document;
    const originalRange = doc.getSelection();
    originalRange.insertText("many", Word.InsertLocation.replace);
    ```

1. Save all your changes to the project.

### Test the add-in

1. [!include[Start server and sideload add-in instructions](../includes/tutorial-word-start-server.md)]

1. If the add-in task pane isn't already open in Word, go to the **Home** tab and choose the **Show Task Pane** button on the ribbon to open it.

1. In the task pane, choose the **Insert Paragraph** button to ensure that there's a paragraph at the start of the document.

1. Within the document, select the phrase "Microsoft 365 subscription". *Be careful not to include the preceding space or following comma in the selection.*

1. Choose the **Insert Abbreviation** button. Note that " (M365)" is added. Note also that at the bottom of the document a new paragraph is added with the entire expanded text because the new string was added to the existing range.

1. Within the document, select the phrase "Microsoft 365". *Be careful not to include the preceding or following space in the selection.*

1. Choose the **Add Version Info** button. Note that "Office 2024, " is inserted between "Office 2021" and "Microsoft 365". Note also that at the bottom of the document a new paragraph is added but it contains only the originally selected text because the new string became a new range rather than being added to the original range.

1. Within the document, select the word "several". *Be careful not to include the preceding or following space in the selection.*

1. Choose the **Change Quantity Term** button. Note that "many" replaces the selected text.

    :::image type="content" source="../images/word-tutorial-text-replace-2.png" alt-text="The results of choosing the add-in buttons Insert Abbreviation, Add Version Info, and Change Quantity Term.":::

## Insert images, HTML, and tables

In this step of the tutorial, you'll learn how to insert images, HTML, and tables into the document.

### Define an image

Complete the following steps to define the image that you'll insert into the document in the next part of this tutorial.

1. In the root of the project, create a new file named **base64Image.js**.

1. Open the file **base64Image.js** and add the following code to specify the Base64-encoded string that represents an image.

    ```js
    export const base64Image =
        "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";
    ```

### Insert an image

1. Open the file **./src/taskpane/taskpane.html**.

1. Locate the `<button>` element for the `replace-text` button, and add the following markup after that line.

    ```html
    <button class="ms-Button" id="insert-image">Insert Image</button><br/><br/>
    ```

1. Open the file **./src/taskpane/taskpane.ts**.

1. Locate the `Office.onReady` function call near the top of the file and add the following code immediately before that line. This code imports the variable that you defined previously in the file **./base64Image.js**.

    ```js
    import { base64Image } from "../../base64Image";
    ```

1. Within the `Office.onReady` function call, locate the line that assigns a click handler to the `replace-text` button, and add the following code after that line.

    ```js
    document.getElementById("insert-image").onclick = () => tryCatch(insertImage);
    ```

1. Add the following function to the end of the file.

    ```js
    async function insertImage() {
        await Word.run(async (context) => {

            // TODO1: Queue commands to insert an image.

            await context.sync();
        });
    }
    ```

1. Within the `insertImage()` function, replace `TODO1` with the following code. Note that this line inserts the Base64-encoded image at the end of the document. (The `Paragraph` object also has an `insertInlinePictureFromBase64` method and other `insert*` methods. See the following "Insert HTML" section for an example.)

    ```js
    context.document.body.insertInlinePictureFromBase64(base64Image, Word.InsertLocation.end);
    ```

### Insert HTML

1. Open the file **./src/taskpane/taskpane.html**.

1. Locate the `<button>` element for the `insert-image` button, and add the following markup after that line.

    ```html
    <button class="ms-Button" id="insert-html">Insert HTML</button><br/><br/>
    ```

1. Open the file **./src/taskpane/taskpane.ts**.

1. Within the `Office.onReady` function call, locate the line that assigns a click handler to the `insert-image` button, and add the following code after that line.

    ```js
    document.getElementById("insert-html").onclick = () => tryCatch(insertHTML);
    ```

1. Add the following function to the end of the file.

    ```js
    async function insertHTML() {
        await Word.run(async (context) => {

            // TODO1: Queue commands to insert a string of HTML.

            await context.sync();
        });
    }
    ```

1. Within the `insertHTML()` function, replace `TODO1` with the following code. Note:

   - The first line adds a blank paragraph to the end of the document.

   - The second line inserts a string of HTML at the end of the paragraph; specifically two paragraphs, one formatted with the Verdana font, the other with the default styling of the Word document. (As you saw in the `insertImage` method earlier, the `context.document.body` object also has the `insert*` methods.)

    ```js
    const blankParagraph = context.document.body.paragraphs.getLast().insertParagraph("", Word.InsertLocation.after);
    blankParagraph.insertHtml('<p style="font-family: verdana;">Inserted HTML.</p><p>Another paragraph</p>', Word.InsertLocation.end);
    ```

### Insert a table

1. Open the file **./src/taskpane/taskpane.html**.

1. Locate the `<button>` element for the `insert-html` button, and add the following markup after that line.

    ```html
    <button class="ms-Button" id="insert-table">Insert Table</button><br/><br/>
    ```

1. Open the file **./src/taskpane/taskpane.ts**.

1. Within the `Office.onReady` function call, locate the line that assigns a click handler to the `insert-html` button, and add the following code after that line.

    ```js
    document.getElementById("insert-table").onclick = () => tryCatch(insertTable);
    ```

1. Add the following function to the end of the file.

    ```js
    async function insertTable() {
        await Word.run(async (context) => {

            // TODO1: Queue commands to get a reference to the paragraph
            //        that will precede the table.

            // TODO2: Queue commands to create a table and populate it with data.

            await context.sync();
        });
    }
    ```

1. Within the `insertTable()` function, replace `TODO1` with the following code. Note that this line uses the `ParagraphCollection.getFirst` method to get a reference to the first paragraph and then uses the `Paragraph.getNext` method to get a reference to the second paragraph.

    ```js
    const secondParagraph = context.document.body.paragraphs.getFirst().getNext();
    ```

1. Within the `insertTable()` function, replace `TODO2` with the following code. Note:

   - The first two parameters of the `insertTable` method specify the number of rows and columns.

   - The third parameter specifies where to insert the table, in this case after the paragraph.

   - The fourth parameter is a two-dimensional array that sets the values of the table cells.

   - The table will have plain default styling, but the `insertTable` method returns a `Table` object with many members, some of which are used to style the table.

    ```js
    const tableData = [
            ["Name", "ID", "Birth City"],
            ["Bob", "434", "Chicago"],
            ["Sue", "719", "Havana"],
        ];
    secondParagraph.insertTable(3, 3, Word.InsertLocation.after, tableData);
    ```

1. Save all your changes to the project.

### Test the add-in

1. [!include[Start server and sideload add-in instructions](../includes/tutorial-word-start-server.md)]

1. If the add-in task pane isn't already open in Word, go to the **Home** tab and choose the **Show Task Pane** button on the ribbon to open it.

1. In the task pane, choose the **Insert Paragraph** button at least three times to ensure that there are a few paragraphs in the document.

1. Choose the **Insert Image** button and note that an image is inserted at the end of the document.

1. Choose the **Insert HTML** button and note that two paragraphs are inserted at the end of the document, and that the first one has the Verdana font.

1. Choose the **Insert Table** button and note that a table is inserted after the second paragraph.

    :::image type="content" source="../images/word-tutorial-insert-image-html-table-2.png" alt-text="The results of choosing the add-in buttons Insert Image, Insert HTML, and Insert Table.":::

## Create and update content controls

In this step of the tutorial, you'll learn how to create Rich Text content controls in the document, and then how to insert and replace content in the controls.

> [!NOTE]
> Before you start this step of the tutorial, we recommend that you create and manipulate Rich Text content controls through the Word UI, so you can be familiar with the controls and their properties. For details, see [Create forms that users complete or print in Word](https://support.microsoft.com/office/040c5cc1-e309-445b-94ac-542f732c8c8b).

### Create a content control

1. Open the file **./src/taskpane/taskpane.html**.

1. Locate the `<button>` element for the `insert-table` button, and add the following markup after that line.

    ```html
    <button class="ms-Button" id="create-content-control">Create Content Control</button><br/><br/>
    ```

1. Open the file **./src/taskpane/taskpane.ts**.

1. Within the `Office.onReady` function call, locate the line that assigns a click handler to the `insert-table` button, and add the following code after that line.

    ```js
    document.getElementById("create-content-control").onclick = () => tryCatch(createContentControl);
    ```

1. Add the following function to the end of the file.

    ```js
    async function createContentControl() {
        await Word.run(async (context) => {

            // TODO1: Queue commands to create a content control.

            await context.sync();
        });
    }
    ```

1. Within the `createContentControl()` function, replace `TODO1` with the following code. Note:

   - This code is intended to wrap the phrase "Microsoft 365" in a content control. It makes a simplifying assumption that the string is present and the user has selected it.

   - The `ContentControl.title` property specifies the visible title of the content control.

   - The `ContentControl.tag` property specifies an tag that can be used to get a reference to a content control using the `ContentControlCollection.getByTag` method, which you'll use in a later function.

   - The `ContentControl.appearance` property specifies the visual look of the control. Using the value "Tags" means that the control will be wrapped in opening and closing tags, and the opening tag will have the content control's title. Other possible values are "BoundingBox" and "None".

   - The `ContentControl.color` property specifies the color of the tags or the border of the bounding box.

    ```js
    const serviceNameRange = context.document.getSelection();
    const serviceNameContentControl = serviceNameRange.insertContentControl();
    serviceNameContentControl.title = "Service Name";
    serviceNameContentControl.tag = "serviceName";
    serviceNameContentControl.appearance = "Tags";
    serviceNameContentControl.color = "blue";
    ```

### Replace the content of the content control

1. Open the file **./src/taskpane/taskpane.html**.

1. Locate the `<button>` element for the `create-content-control` button, and add the following markup after that line.

    ```html
    <button class="ms-Button" id="replace-content-in-control">Rename Service</button><br/><br/>
    ```

1. Open the file **./src/taskpane/taskpane.ts**.

1. Within the `Office.onReady` function call, locate the line that assigns a click handler to the `create-content-control` button, and add the following code after that line.

    ```js
    document.getElementById("replace-content-in-control").onclick = () => tryCatch(replaceContentInControl);
    ```

1. Add the following function to the end of the file.

    ```js
    async function replaceContentInControl() {
        await Word.run(async (context) => {

            // TODO1: Queue commands to replace the text in the Service Name
            //        content control.

            await context.sync();
        });
    }
    ```

1. Within the `replaceContentInControl()` function, replace `TODO1` with the following code. Note:

    - The `ContentControlCollection.getByTag` method returns a `ContentControlCollection` of all content controls of the specified tag. We use `getFirst` to get a reference to the desired control.

    ```js
    const serviceNameContentControl = context.document.contentControls.getByTag("serviceName").getFirst();
    serviceNameContentControl.insertText("Fabrikam Online Productivity Suite", Word.InsertLocation.replace);
    ```

1. Save all your changes to the project.

### Test the add-in

1. [!include[Start server and sideload add-in instructions](../includes/tutorial-word-start-server.md)]

1. If the add-in task pane isn't already open in Word, go to the **Home** tab and choose the **Show Task Pane** button on the ribbon to open it.

1. In the task pane, choose the **Insert Paragraph** button to ensure that there's a paragraph with "Microsoft 365" at the top of the document.

1. In the document, select the text "Microsoft 365" and then choose the **Create Content Control** button. Note that the phrase is wrapped in tags labelled "Service Name".

1. Choose the **Rename Service** button and note that the text of the content control changes to "Fabrikam Online Productivity Suite".

    :::image type="content" source="../images/word-tutorial-content-control-2.png" alt-text="The results of choosing the add-in buttons Create Content Control and Rename Service.":::

## Next steps

In this tutorial, you've created a Word task pane add-in that inserts and replaces text, images, and other content in a Word document. To learn more about building Word add-ins, continue to the following article.

> [!div class="nextstepaction"]
> [Word add-ins overview](../word/word-add-ins-programming-overview.md)

## Code samples

- [Completed Word add-in tutorial](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/tutorials/word-tutorial): The result of completing this tutorial.

## See also

- [Office Add-ins platform overview](../overview/office-add-ins.md)
- [Develop Office Add-ins](../develop/develop-overview.md)
