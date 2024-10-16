---
title: PowerPoint add-ins
description: Learn how to use PowerPoint add-ins to build engaging solutions for presentations across platforms including Windows, iPad, Mac, and in a browser.
ms.date: 10/16/2024
ms.topic: overview
ms.custom: scenarios:getting-started
ms.localizationpriority: high
---

# PowerPoint add-ins

You can use PowerPoint add-ins to build engaging solutions for your users' presentations across platforms including Windows, iPad, Mac, and in a browser. You can create two types of PowerPoint add-ins:

- Use **content add-ins** to add dynamic HTML5 content to your presentations. For example, see the [LucidChart Diagrams for PowerPoint](https://appsource.microsoft.com/product/office/wa104380117) add-in, which you can use to inject an interactive diagram from LucidChart into your deck. To create your own content add-in, you can start with [Build your first PowerPoint content add-in](../quickstarts/powerpoint-quickstart-content.md).

- Use **task pane add-ins** to bring in reference information or insert data into the presentation via a service. For example, see the [Pexels - Free Stock Photos](https://appsource.microsoft.com/product/office/wa104379997) add-in, which you can use to add professional photos to your presentation. To create your own task pane add-in, you can start with [Build your first PowerPoint task pane add-in](../quickstarts/powerpoint-quickstart-yo.md).

## PowerPoint add-in scenarios

The code examples in this article demonstrate some basic tasks for developing add-ins for PowerPoint. Please note the following:

- To display information, these examples use the `app.showNotification` function, which is included in the Visual Studio Office Add-ins project templates. If you aren't using Visual Studio to develop your add-in, you'll need replace the `showNotification` function with your own code.

- Several of these examples also use a `Globals` object that's declared beyond the scope of these functions as:

   `let Globals = {activeViewHandler:0, firstSlideId:0};`

- To use these examples, your add-in project must [reference Office.js v1.1 library or later](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).

## Detect the presentation's active view and handle the ActiveViewChanged event

If you're building a content add-in, you'll need to get the presentation's active view and handle the `ActiveViewChanged` event, as part of your `Office.initialize` handler.

> [!NOTE]
> In PowerPoint on the web, the [Document.ActiveViewChanged](/javascript/api/office/office.document) event will never fire as Slide Show mode is treated as a new session. In this case, the add-in must fetch the active view on load, as shown in the following code sample.

Note the following about the code sample:

- The  `getActiveFileView` function calls the [Document.getActiveViewAsync](/javascript/api/office/office.document#office-office-document-getactiveviewasync-member(1)) method to return whether the presentation's current view is "edit" (any of the views in which you can edit slides, such as **Normal** or **Outline View**) or "read" (**Slide Show** or **Reading View**).

- The  `registerActiveViewChanged` function calls the [addHandlerAsync](/javascript/api/office/office.document#office-office-document-addhandlerasync-member(1)) method to register a handler for the [Document.ActiveViewChanged](/javascript/api/office/office.document) event.

```js
// General Office.initialize function. Called after the add-in loads and Office JS is initialized.
Office.initialize = function(){

    // Get whether the current view is edit or read.
    const currentView = getActiveFileView();

    // Register for the active view changed handler.
    registerActiveViewChanged();

    // Render the content based off of the currentView.
    //....
}

function getActiveFileView()
{
    Office.context.document.getActiveViewAsync(function (asyncResult) {
        if (asyncResult.status === "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        } else {
            app.showNotification(asyncResult.value);
        }
    });

}

function registerActiveViewChanged() {
    Globals.activeViewHandler = function (args) {
        app.showNotification(JSON.stringify(args));
    }

    Office.context.document.addHandlerAsync(Office.EventType.ActiveViewChanged, Globals.activeViewHandler,
        function (asyncResult) {
            if (asyncResult.status === "failed") {
                app.showNotification("Action failed with error: " + asyncResult.error.message);
            } else {
                app.showNotification(asyncResult.status);
            }
        });
}
```

## Navigate to a particular slide in the presentation

In the following code sample, the `getSelectedRange` function calls the [Document.getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) method to get the JSON object returned by `asyncResult.value`, which contains an array named `slides`. The `slides` array contains the IDs, titles, and indexes of selected range of slides (or of the current slide, if multiple slides aren't selected). It also saves the ID of the first slide in the selected range to a global variable.

```js
function getSelectedRange() {
    // Gets the ID, title, and index of the current slide (or selected slides) and store the first slide ID. */
    Globals.firstSlideId = 0;

    Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function (asyncResult) {
        if (asyncResult.status === "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        } else {
            Globals.firstSlideId = asyncResult.value.slides[0].id;
            app.showNotification(JSON.stringify(asyncResult.value));
        }
    });
}
```

In the following code sample, the `goToFirstSlide` function calls the [Document.goToByIdAsync](/javascript/api/office/office.document#office-office-document-gotobyidasync-member(1)) method to navigate to the first slide that was identified by the `getSelectedRange` function shown previously.

```js
function goToFirstSlide() {
    Office.context.document.goToByIdAsync(Globals.firstSlideId, Office.GoToType.Slide, function (asyncResult) {
        if (asyncResult.status === "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        } else {
            app.showNotification("Navigation successful");
        }
    });
}
```

## Navigate between slides in the presentation

In the following code sample, the `goToSlideByIndex` function calls the `Document.goToByIdAsync` method to navigate to the next slide in the presentation.

```js
function goToSlideByIndex() {
    const goToFirst = Office.Index.First;
    const goToLast = Office.Index.Last;
    const goToPrevious = Office.Index.Previous;
    const goToNext = Office.Index.Next;

    Office.context.document.goToByIdAsync(goToNext, Office.GoToType.Index, function (asyncResult) {
        if (asyncResult.status === "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        } else {
            app.showNotification("Navigation successful");
        }
    });
}
```

## Get the URL of the presentation

In the following code sample, the  `getFileUrl` function calls the [Document.getFileProperties](/javascript/api/office/office.document#office-office-document-getfilepropertiesasync-member(1)) method to get the URL of the presentation file.

```js
function getFileUrl() {
    // Gets the URL of the current file.
    Office.context.document.getFilePropertiesAsync(function (asyncResult) {
        const fileUrl = asyncResult.value.url;
        if (fileUrl === "") {
            app.showNotification("The file hasn't been saved yet. Save the file and try again.");
        } else {
            app.showNotification(fileUrl);
        }
    });
}
```

## Create a presentation

Your add-in can create a new presentation, separate from the PowerPoint instance in which the add-in is currently running. The PowerPoint namespace has the `createPresentation` method for this purpose. When this method is called, the new presentation is immediately opened and displayed in a new instance of PowerPoint. Your add-in remains open and running with the previous presentation.

```js
PowerPoint.createPresentation();
```

The `createPresentation` method can also create a copy of an existing presentation. The method accepts a Base64-encoded string representation of an .pptx file as an optional parameter. The resulting presentation will be a copy of that file, assuming the string argument is a valid .pptx file. The [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) class can be used to convert a file into the required Base64-encoded string, as demonstrated in the following example.

```js
const myFile = document.getElementById("file");
const reader = new FileReader();

reader.onload = function (event) {
    // Strip off the metadata before the Base64-encoded string.
    const startIndex = reader.result.toString().indexOf("base64,");
    const copyBase64 = reader.result.toString().substr(startIndex + 7);

    PowerPoint.createPresentation(copyBase64);
};

// Read in the file as a data URL so we can parse the Base64-encoded string.
reader.readAsDataURL(myFile.files[0]);
```

## See also

- [Developing Office Add-ins](../develop/develop-overview.md)
- [Learn about the Microsoft 365 Developer Program](https://aka.ms/m365devprogram)
- PowerPoint quick starts
  - [Build your first PowerPoint content add-in](../quickstarts/powerpoint-quickstart-content.md)
  - [Build your first PowerPoint task pane add-in](../quickstarts/powerpoint-quickstart-yo.md)
- [PowerPoint Code Samples](https://developer.microsoft.com/microsoft-365/gallery/?filterBy=Samples,PowerPoint)
- [How to save add-in state and settings per document for content and task pane add-ins](../develop/persisting-add-in-state-and-settings.md#how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins)
- [Read and write data to the active selection in a document or spreadsheet](../develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)
- [Get the whole document from an add-in for PowerPoint or Word](../powerpoint/get-the-whole-document-from-an-add-in-for-powerpoint.md)
- [Use document themes in your PowerPoint add-ins](use-document-themes-in-your-powerpoint-add-ins.md)
