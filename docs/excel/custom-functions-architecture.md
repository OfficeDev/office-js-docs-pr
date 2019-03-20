---
ms.date: 03/20/2019
description: Learn about Excel custom functions' runtime. 
title: Custom functions architecture (preview)
localization_priority: Priority
---
# Custom functions architecture

Custom function use a new JavaScript runtime that differs from the browser-based JavaScript engine which powers most other parts of your add-in.

## Custom functions runtime

The custom functions runtime prioritizes executing calculations, so functions run run smoothly when Excel recalculates. The custom functions runtime allows access to APIs which make requesting external data and exchanging data over a persistent connection over a server possible.

Because functions are limited to actions performed in a workbook, the custom functions runtime doesn't support UI elements nor does it provide access to the Document Object Model (DOM). You should think of custom functions as a separate component of your add-in, separate from your add-in's task pane or other UI elements.

>[!NOTE] 
> When first installing a custom function, a task pane is shown. This pane is necessary to install the custom functions runtime, but shouldn't serve as your main task pane for your add-in.

## Browser-based engine

You may be used to the common browser-based engine which powers most add-ins if you have developed an add-in before. This engine allows access to the Office.js APIs. Any of the Excel APIs (such as APIs which allow you to manipulate Excel tables, charts, etc) you may have used when developing an Excel add-in run on this browser-based engine, but aren't directly accessible from the custom functions runtime. 

The browser-based engine is also what supports the task pane and UI elements of your add-in. Keep this in mind while developing an add-in which utilizes both a task pane and a custom function, especially if those two parts of your add-in need to communicate.

## Communicate between engines

Both the browser-based engine and custom functions runtime have access to new APIs in the `OfficeRuntime` namespace. Use these APIs to facilitate passing information between your function and the UI elements of your add-in. `OfficeRuntime` APIs offer a storage location (`AsyncStorage`) and a new dialog API which is compatible with custom functions.

## Sample scenarios

`AsyncStorage` is helpful in almost every situation in which you wish to use Office.js APIs with a custom function. One common scenario is authenticating users through the task pane before they are given access to a custom function. You'll find a code sample of this authentication scenario in this [Github repository](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) dedicated to patterns and practices.

For more general information about `AsyncStorage`, see [Custom functions runtime](./custom-functions-runtime.md).

## See also

* [Runtime for Excel custom functions](custom-functions-runtime.md)
* [Custom functions best practices](custom-functions-best-practices.md)
* [Custom functions metadata](custom-functions-json.md)
* [Custom functions changelog](custom-functions-changelog.md)
* [Excel custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md)
