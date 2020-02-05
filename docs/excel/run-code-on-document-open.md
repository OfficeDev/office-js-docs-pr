---
title: Run a custom function when a document opens
description: Write custom functions that run upon a document opening. 
ms.date: 02/05/2020
localization_priority: Normal
---

# Run a custom function when a document opens

By default, custom functions do not run automatically when you open a document, but only when you choose to run them. However, a custom function can automatically when you open a document if you configure your add-in's start up behavior.

This is useful for particular scenarios, such as an add-in which needs to load a custom function immediately or when you wish to register a set of event handlers.

To do this, your add-in's manifest and task pane HTML file must be properly configured as shown in [Share data and events between Excel custom functions and task pane tutorial](TODO LINK).

## Code sample

The following code samples illustrate how to configure the start up behavior for your add-in, which could include running a custom function. These code samples also assume you have already set up your add-in's manifest and task pane HTML file so custom functions can utilize Office.js APIs.

[!NOTE] The `setStartupBehavior` method is asynchronous.

The following code would set the add-in to load immediately the next time the document is opened.

TODO see if this works and if I need to put some cf code somewhere in this
```JavaScript
 Office.addin.setStartupBehavior(Office.StartupBehavior.load);
```

To set the add-in to not load when the document is next opened, pass `none` as the parameter:
```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.none);
```

To determine what the current startup behavior is, run the following function, which returns an Office.StartupBehavior object.

```JavaScript
var behavior = await Office.addin.getStartupBehavior();
```

## Next steps
TODO 

## See also
* [Call Excel APIs from a custom function](TODO link)