---
title: Run code in your Office Add-in when the document opens
description: Learn how to run code in your Office Add-in add-in when the document opens.
ms.topic: how-to
ms.date: 09/22/2025
ms.localizationpriority: medium
---

# Run code in your Office Add-in when the document opens

[!include[Shared runtime requirements](../includes/shared-runtime-requirements-note.md)]

You can configure your Office Add-in to load and run code as soon as the document is opened. This is useful if you need to register event handlers, pre-load data for the task pane, synchronize UI, or perform other tasks before the add-in is visible.

> [!NOTE]
> The configuration is implemented with a method that your code calls at runtime. This means that the add-in *won't* run the *first time* a user opens the document. The add-in must be opened manually for the first time on any document. After the method runs, either in [Office.initialize](/javascript/api/office#office-office-initialize-function(1)), [Office.onReady](/javascript/api/office#office-office-onready-function(1)), or because the user takes a code path that runs it; then whenever the document is reopened, the add-in loads immediately and any code in the `Office.initialize` or `Office.onReady` method runs.

[!include[Shared runtime note](../includes/note-requires-shared-runtime.md)]

## Configure your add-in to load when the document opens

The following code configures your add-in to load and start running when the document is opened.

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.load);
```

> [!NOTE]
> The `setStartupBehavior` method is asynchronous.

## Place startup code in Office.initialize or Office.onReady

When your add-in is configured to load on document open, it will run immediately. The `Office.initialize` event handler will be called. Place your startup code in the `Office.initialize` or `Office.onReady` event handler.

The following Excel add-in code shows how to register an event handler for change events from the active worksheet. If you configure your add-in to load on document open, this code will register the event handler when the document is opened. You can handle change events before the task pane is opened.

```JavaScript
// This is called as soon as the document opens.
// Put your startup code here.
Office.initialize = () => {
  // Add the event handler.
  Excel.run(async context => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.onChanged.add(onChange);

    await context.sync();
    console.log("A handler has been registered for the onChanged event.");
  });
};

/**
 * Handle the changed event from the worksheet.
 *
 * @param event The event information from Excel
 */
async function onChange(event) {
    await Excel.run(async (context) => {    
        await context.sync();
        console.log("Change type of event: " + event.changeType);
        console.log("Address of event: " + event.address);
        console.log("Source of event: " + event.source);
  });
}
```

The following PowerPoint add-in code shows how to register an event handler for selection change events from the PowerPoint document. If you configure your add-in to load on document open, this code will register the event handler when the document is opened. You can handle change events before the task pane is opened.

```JavaScript
// This is called as soon as the document opens.
// Put your startup code here.
Office.onReady(info => {
  if (info.host === Office.HostType.PowerPoint) {
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onChange);
    console.log("A handler has been registered for the onChanged event.");
  }
});

/**
 * Handle the changed event from the PowerPoint document.
 *
 * @param event The event information from PowerPoint
 */
async function onChange(event) {
  console.log("Change type of event: " + event.type);
}
```

## Configure your add-in for no load behavior on document open

There may be scenarios when you want to turn off the "run on document open" behavior. The following code configures your add-in not to start when the document is opened. Instead, it will start when the user engages it in some way, such as choosing a ribbon button or opening the task pane. This code has no effect if the method hasn't previously been called on the current document, with `Office.StartupBehavior.load` as the parameter.

> [!NOTE]
> If add-in calls the method, with `Office.StartupBehavior.load` as the parameter, in `Office.initialize` or `Office.onReady`, the behavior is turned on again. So, in this scenario, turning it off only applies to the *next* time the document is opened, not *all* subsequent openings. 

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.none);
```

## Get the current load behavior

There may be scenarios in which your add-in needs to know if it's configured to start automatically the next time the current document is opened. To determine what the current startup behavior is, run the following method, which returns an [Office.StartupBehavior](/javascript/api/office/office.startupbehavior) value.

```JavaScript
let behavior = await Office.addin.getStartupBehavior();
```

## See also

- [Configure your Office Add-in to use a shared runtime](configure-your-add-in-to-use-a-shared-runtime.md)
- [Share data and events between Excel custom functions and task pane tutorial](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [Work with Events using the Excel JavaScript API](../excel/excel-add-ins-events.md)
- [Runtimes in Office Add-ins](../testing/runtimes.md)
- [Managing trust options for Office Add-ins](../publish/manage-trust-options.md)
