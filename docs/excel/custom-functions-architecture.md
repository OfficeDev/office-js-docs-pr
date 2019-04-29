---
ms.date: 03/20/2019
description: Learn about Excel custom functions' runtime. 
title: Custom functions architecture (preview)
localization_priority: Priority
---
# Custom functions architecture

 Custom functions are with their own unique runtime that prioritizes execution of calculations. This article will cover the differences between the custom functions runtime and the browser-based JavaScript engine which powers most other parts of your add-in.

## Custom functions runtime

An Office Web Add-in can interact with the user as a task pane, or a content pane, and can include commands and custom functions. All of these parts run in a browser engine runtime except for custom functions. Custom functions run in a separate custom functions runtime to optimize for calculation speed.

Note that if you're using the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) to generate your project, the custom functions runtime will load through the custom-functions.js script file referenced in the **functions.html** file. The **functions.html** serves only to load the runtime and shouldn't be used as the task pane for your add-in.

The following table highlights the differences between the custom functions runtime and the browser engine runtime:

| Custom functions runtime 	| Browser engine runtime 	|
|------------------------------------------------------------------	|--------------------------------------------------------------------------------------------------------------	|
| Supports returning a value from a cell 	| Supports Office.js APIs and UI elements 	|
| Does not have `localStorage` object, instead uses `AsyncStorage` 	| Has `localStorage` object, can optionally use `AsyncStorage` object 	|
| Does not support interacting with the DOM, or loading libraries that depend on the DOM such as jQuery.	| Supports interacting with the DOM and loading libraries that depend on the DOM. |


## Browser engine runtime

The task pane, content add-in, and commands run in a browser engine runtime.

The browser engine runtime supports the Office.js APIs. Keep in mind that any of the Excel APIs, such as APIs which allow you to manipulate Excel tables, run on the browser engine runtime, but aren't directly accessible from the custom functions runtime.

## Communicate between runtimes

Your custom functions code cannot directly interact with code in other parts of your web add-in, like the task pane because they are in different runtimes. But in some scenarios you may need to share data, such as passing a token.

`AsyncStorage` can be used to store data from your custom functions and get data from your task pane code. For more information about storing and sharing data, see [Saving and sharing state](custom-functions-overview.md#saving-and-sharing-state).

You can see a code sample using `AsyncStorage` in this [Github repository](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) dedicated to patterns and practices.
For more general information about `AsyncStorage`, see [Custom functions runtime](./custom-functions-runtime.md).

`AsyncStorage` can also be useful for authentication. For more information, see [Custom functions authentication](custom-functions-authentication.md).

## See also

* [Runtime for Excel custom functions](custom-functions-runtime.md)
* [Custom functions best practices](custom-functions-best-practices.md)
* [Custom functions metadata](custom-functions-json.md)
* [Custom functions changelog](custom-functions-changelog.md)
* [Excel custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md)
