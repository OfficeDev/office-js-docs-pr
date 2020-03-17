---
title: Run code in your Excel add-in when the document opens (preview)
description: 'Run code in your Excel add-in when the document opens.'
ms.date: 02/20/2020
localization_priority: Normal
---

# Run code in your Excel add-in when the document opens (preview)

[!include[Running custom functions in browser runtime note](../includes/excel-shared-runtime-preview-note.md)]

You can configure your Excel add-in to load and run code as soon as the document is opened. This is useful if you need to register event handlers, preload data for the task pane, synchronize UI, or perform other tasks before the add-in is visible.

[!include[Excel shared runtime note](../includes/note-requires-shared-runtime.md)]

## Configure your add-in to load when the document opens

The following code configures your add-in to load and start running when the document is opened.

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.load);
```

> [!NOTE]
> The `setStartupBehavior` method is asynchronous.

## Configure your add-in for no load behavior on document open

The following code configures your add-in not to start when the document is opened. Instead it will start when the user engages it in some way (such as choosing a ribbon button, or opening the task pane.)

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.none);
```

## Get the current load behavior

To determine what the current startup behavior is, run the following function, which returns an Office.StartupBehavior object.

```JavaScript
let behavior = await Office.addin.getStartupBehavior();
```

## How to run code when the document opens

When your add-in is configured to load on document open, it will run immediately. The `Office.initialize` event handler will be called. Place your startup code in the `Office.initialize` event handler.

The following code shows how to register an event handler for change events from the active worksheet. If you configure your add-in to load on document open, this code will register the event handler when the document is opened. You can handle change events before the task pane is opened.


```JavaScript
//This is called as soon as the document opens.
//Put your startup code here.
Office.initialize = () => {
  // Add the event handler
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
  return Excel.run(function(context) {
    return context.sync().then(function() {
      console.log("Change type of event: " + event.changeType);
      console.log("Address of event: " + event.address);
      console.log("Source of event: " + event.source);
    });
  });
}

```

## See also

- [Share data and events between Excel custom functions and task pane tutorial](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)