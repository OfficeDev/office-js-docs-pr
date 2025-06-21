---
title: PowerPoint add-ins
description: Learn how to use PowerPoint add-ins to build engaging solutions for presentations across platforms including Windows, iPad, Mac, and in a browser.
ms.date: 06/05/2025
ms.topic: overview
ms.custom: scenarios:getting-started
ms.localizationpriority: high
---

# PowerPoint add-ins

You can use PowerPoint add-ins to build engaging solutions for your users' presentations across platforms including Windows, iPad, Mac, and in a browser. You can create two types of PowerPoint add-ins:

- Use **task pane add-ins** to bring in reference information or insert data into the presentation via a service. For example, see the [Pexels - Free Stock Photos](https://appsource.microsoft.com/product/office/wa104379997) add-in, which you can use to add professional photos to your presentation. To create your own task pane add-in, you can start with [Build your first PowerPoint task pane add-in](../quickstarts/powerpoint-quickstart-yo.md).

- Use **content add-ins** to add dynamic HTML5 content to your presentations. For example, see the [LucidChart Diagrams for PowerPoint](https://appsource.microsoft.com/product/office/wa104380117) add-in, which injects interactive diagrams from LucidChart into your deck. To create your own content add-in, start with [Build your first PowerPoint content add-in](../quickstarts/powerpoint-quickstart-content.md).

## PowerPoint add-in scenarios

The code examples in this article demonstrate some basic tasks that can be useful when developing add-ins for PowerPoint.

## Add a new slide then navigate to it

In the following code sample, the `addAndNavigateToNewSlide` function calls the [SlideCollection.add](/javascript/api/powerpoint/powerpoint.slidecollection#powerpoint-powerpoint-slidecollection-add-member(1)) method to add a new slide to the presentation. The function then calls the [Presentation.setSelectedSlides](/javascript/api/powerpoint/powerpoint.presentation#powerpoint-powerpoint-presentation-setselectedslides-member(1)) method to navigate to the new slide.

```js
async function addAndNavigateToNewSlide() {
  // Adds a new slide then navigates to it.
  await PowerPoint.run(async (context) => {
    const slideCountResult = context.presentation.slides.getCount();
    context.presentation.slides.add();
    await context.sync();

    const newSlide = context.presentation.slides.getItemAt(slideCountResult.value);
    newSlide.load("id");
    await context.sync();

    console.log(`Added slide - ID: ${newSlide.id}`);

    // Navigate to the new slide.
    context.presentation.setSelectedSlides([newSlide.id]);
    await context.sync();
  });
}
```

## Navigate to a particular slide in the presentation

In the following code sample, the `getSelectedSlides` function calls the [Presentation.getSelectedSlides](/javascript/api/powerpoint/powerpoint.presentation#powerpoint-powerpoint-presentation-getselectedslides-member(1)) method to get the selected slides then logs their IDs. The function can then act on the current slide (or first slide from the selection).

```js
async function getSelectedSlides() {
  // Gets the ID of the current slide (or selected slides).
  await PowerPoint.run(async (context) => {
    const selectedSlides = context.presentation.getSelectedSlides();
    selectedSlides.load("items/id");
    await context.sync();

    if (selectedSlides.items.length === 0) {
      console.warn("No slides were selected.");
      return;
    }

    console.log("IDs of selected slides:");
    selectedSlides.items.forEach(item => {
      console.log(item.id);
    });

    // Navigate to first selected slide.
    const currentSlide = selectedSlides.items[0];
    console.log(`Navigating to slide with ID ${currentSlide.id} ...`);
    context.presentation.setSelectedSlides([currentSlide.id]);

    // Perform actions on current slide...
  });
}
```

## Navigate between slides in the presentation

In the following code sample, the `goToSlideByIndex` function calls the `Presentation.setSelectedSlides` method to navigate to the first slide in the presentation, which has the index 0. The maximum slide index you can navigate to in this sample is `slideCountResult.value - 1`.

```js
async function goToSlideByIndex() {
  await PowerPoint.run(async (context) => {
    // Gets the number of slides in the presentation.
    const slideCountResult = context.presentation.slides.getCount();
    await context.sync();

    if (slideCountResult.value === 0) {
      console.warn("There are no slides.");
      return;
    }

    const slide = context.presentation.slides.getItemAt(0); // First slide
    //const slide = context.presentation.slides.getItemAt(slideCountResult.value - 1); // Last slide
    slide.load("id");
    await context.sync();

    console.log(`Slide ID: ${slide.id}`);

    // Navigate to the slide.
    context.presentation.setSelectedSlides([slide.id]);
    await context.sync();
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
      console.warn("The file hasn't been saved yet. Save the file and try again.");
    } else {
      console.log(`File URL: ${fileUrl}`);
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
const myFile = document.getElementById("file") as HTMLInputElement;
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

To see a full code sample that includes an HTML implementation, see [Create presentation](https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/powerpoint/document/create-presentation.yaml).

## Detect the presentation's active view and handle the ActiveViewChanged event

If you're building a [content add-in](../design/content-add-ins.md), you'll need to get the presentation's active view and handle the [Document.ActiveViewChanged](/javascript/api/office/office.eventtype#fields) event as part of your [Office.onReady](/javascript/api/office#office-office-onready-function(1)) call.

> [!NOTE]
> In PowerPoint on the web, the `Document.ActiveViewChanged` event will never fire because **Slide Show** mode is treated as a new session. In this case, the add-in must fetch the active view on load, as shown in the following code sample.

Note the following about the code sample:

- The `getActiveFileView` function calls the [Document.getActiveViewAsync](/javascript/api/office/office.document#office-office-document-getactiveviewasync-member(1)) method to return whether the presentation's current view is "edit" (any of the view where you can edit slides, such as **Normal**, **Slide Sorter**, or **Outline**) or "read" (**Slide Show** or **Reading View**), represented by the [ActiveView](/javascript/api/office/office.activeview) enum.
- The `registerActiveViewChanged` function calls the [Document.addHandlerAsync](/javascript/api/office/office.document#office-office-document-addhandlerasync-member(1)) method to register a handler for the `Document.ActiveViewChanged` event.
- To display information, this example uses the `showNotification` function, which is included in the Visual Studio Office Add-ins project templates. If you aren't using Visual Studio to develop your add-in, you'll need to replace the `showNotification` function with your own code.

```js
// General Office.onReady function. Called after the add-in loads and Office JS is initialized.
Office.onReady(() => {
  // Get whether the current view is edit or read.
  const currentView = getActiveFileView();

  // Register the active view changed handler.
  registerActiveViewChanged();

  // Render the content based off of the current view.
  if (currentView === Office.ActiveView.Read) {
      // Handle read view.
      console.log('Current view is read.');
      // You can add any specific logic for the read view here.
  } else {
      // Handle edit view.
      console.log('Current view is edit.');
      // You can add any specific logic for the edit view here.
  }
});

// Gets the active file view.
function getActiveFileView() {
    console.log('Getting active file view...');
    Office.context.document.getActiveViewAsync(function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            console.log('Active view:', result.value);
            return result.value;
        } else {
            console.error('Error getting active view:', result.error.message);
            showNotification('Error:', result.error.message);
            return null;
        }
    });
}

// Registers the ActiveViewChanged event.
function registerActiveViewChanged() {
    console.log('Registering ActiveViewChanged event handler...');
    Office.context.document.addHandlerAsync(
        Office.EventType.ActiveViewChanged,
        activeViewHandler,
        function (result) {
            if (result.status === Office.AsyncResultStatus.Failed) {
                console.error('Failed to register active view changed handler:', result.error.message);
                showNotification('Error:', result.error.message);
            } else {
                console.log('Active view changed handler registered successfully.');
            }
        });
}

// ActiveViewChanged event handler.
function activeViewHandler(eventArgs) {
    console.log('Active view changed:', JSON.stringify(eventArgs));
    showNotification('Active view changed', `The active view has changed to: ${eventArgs.activeView}`);
    // You can add logic here based on the new active view.
}
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
- [Get the whole document from an add-in for PowerPoint or Word](../develop/get-the-whole-document-from-an-add-in-for-powerpoint-or-word.md)
- [Use document themes in your PowerPoint add-ins](use-document-themes-in-your-powerpoint-add-ins.md)
