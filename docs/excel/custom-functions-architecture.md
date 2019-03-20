---
ms.date: 03/20/2019
description: Learn about Excel custom functions' runtime. 
title: Custom functions architecture (preview)
localization_priority: Priority
---
# Custom functions architecture

 Custom functions are with their own unique runtime that prioritizes execution of calculations. This article will cover the differences between the custom functions runtime and the browser-based JavaScript engine which powers most other parts of your add-in.

## Custom functions runtime

Because of their unique runtime, think of custom functions as another component of your total add-in. Your add-in will also likely include elements from the browser engine runtime, like a task pane.

The following table highlights the differences between the custom functions runtime and the browser engine runtime:

| Custom functions runtime 	| Browser engine runtime 	|
|------------------------------------------------------------------	|--------------------------------------------------------------------------------------------------------------	|
| Supports returning a value from a cell 	| Supports Office.js APIs and UI elements 	|
| Does not have `localStorage` object, instead uses `AsyncStorage` 	| Has `localStorage` object, can optionally use `AsyncStorage` object 	|
| Provides access to `OfficeRuntime` objects 	| Provides access to `OfficeRuntime` objects 	|
| Allows you to make web requests 	| Allows you to access the Document Object Model (DOM) and supports libraries that use the DOM, such as jQuery 	|

>[!NOTE]
> When first installing a custom function, a task pane is shown. This pane is necessary to install the custom functions runtime, but shouldn't serve as the task pane for your add-in.

## Browser engine runtime

The task pane and other components that are not custom functions run in a browser engine runtime.

The browser engine runtime enables access to the Office.js APIs. Keep in mind that any of the Excel APIs, such as APIs which allow you to manipulate Excel tables, run on the browser engine runtime, but aren't directly accessible from the custom functions runtime.

## Communicate between runtimes

Both the browser engine runtime and the custom functions runtime have access to `OfficeRuntime.AsyncStorage` and `OfficeRuntime.Dialog` APIs.

`AsyncStorage` can be used to store data from your custom functions and get data from your task pane code.

You'll find a code sample using `AsyncStorage` in this [Github repository](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) dedicated to patterns and practices.
For more general information about `AsyncStorage`, see [Custom functions runtime](./custom-functions-runtime.md).

`AsyncStorage` can also be useful for authentication. For more information, see [Custom functions authentication](custom-functions-authentication.md).

## See also

* [Runtime for Excel custom functions](custom-functions-runtime.md)
* [Custom functions best practices](custom-functions-best-practices.md)
* [Custom functions metadata](custom-functions-json.md)
* [Custom functions changelog](custom-functions-changelog.md)
* [Excel custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md)
