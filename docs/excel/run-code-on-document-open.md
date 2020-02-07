---
title: Run code in your Excel add-in when the document opens (preview)
description: Run code in your Excel add-in when the document opens. 
ms.date: 02/06/2020
localization_priority: Normal
---

# Run code in your Excel add-in when the document opens (preview)

[!include[Running custom functions in browser runtime note](../includes/excel-shared-runtime-preview-note.md)]

You can configure your Excel add-in to load and run code as soon as the document is opened. This is useful if you need to register event handlers, preload data for the task pane, synchronize UI, or perform other tasks before the add-in is visible.

[!include[Excel shared runtime note](../includes/note-requires-shared-runtime.md)]

## Code samples for configuring load behavior

The following code samples illustrate how to configure the start up behavior for your add-in, which could include running a custom function. These code samples also assume you have already modified your add-in's manifest and task pane HTML file so custom functions can utilize Office.js APIs.

### Set add-in to start on document open

The following code sets the add-in to load immediately the next time the document is opened.

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.load);
```

> [!NOTE]
> The `setStartupBehavior` method is asynchronous.

### Set add-in not to start on document open

To set the add-in to not load when the document is next opened, pass `none` as the parameter:

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.none);
```

### Get the current load behavior

To determine what the current startup behavior is, run the following function, which returns an Office.StartupBehavior object.

```JavaScript
let behavior = await Office.addin.getStartupBehavior();
```

## Code sample for running when document loads

When your add-in is configured to load on document open, it will run immediately. The `Office.initialize` event handler will be called. Place your startup code in the `Office.initialize` event handler.

An example scenario of running code when the document opens is to display your task pane immediately. The following code shows how to show the task pane as soon as the document is opened.

```JavaScript
//This is called as soon as the document opens.
//Put your startup code here.
Office.initialize = () => {
  // Display the task pane
  SetRuntimeVisibleHelper(true);
};

//Display or hide the task pane based on visible parameter
function SetRuntimeVisibleHelper = (visible) => {
  let p;
  if (visible) {
    p = Office.addin.showAsTaskpane();
  }
  else {
    p = Office.addin.hide();
  }
  return p.then(() => {
    return visible;
  })
  .catch((error) => {
    return error.code;
  });
};
```

## See also

- [Share data and events between Excel custom functions and task pane tutorial](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)