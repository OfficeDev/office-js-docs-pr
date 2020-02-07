---
ms.date: 02/01/2020
description: Learn how to run custom functions, ribbon buttons, and task pane code in a single browser runtime to coordinate scenarios across your add-in.
title: Run your add-in code in a single browser runtime (preview)
localization_priority: Priority
---

# Overview: Run your add-in code in a single browser runtime (preview)

[!include[Running custom functions in browser runtime note](../includes/excel-shared-runtime-preview-note.md)]

When running Excel on Windows or Mac, ribbon buttons, custom functions, and the task pane code run in separate JavaScript runtime environments. You can configure your Excel add-in to run code in a single browser runtime. This is also known as a shared runtime since the same JavaScript runtime is used for different parts of your add-in. 

Configuring a shared runtime enables the following scenarios:
- Custom functions can take full advantage of the browser runtime to get full CORS support, use the DOM, and enable additional web scenarios.
- Custom functions can call Excel JS APIs to read spreadsheet document data.
- You can run code as soon as the document is opened.
- You can continue running code after the task pane is closed.

When you run custom functions in the same browser runtime as the task pane, it will run in a browser instance on different platforms as follows:

- Excel on Windows: Code for custom functions and the task pane run in an IE or Edge instance.
- Excel on Mac: Code for custom functions and the task pane run in a Safari instance.
- Excel online: Code for custom functions and the task pane run in the same browser instance (Chrome, Edge, Safari, FireFox, etc...)

Additionally, any buttons that your Excel add-in displays on the ribbon will run in the same browser runtime. The following image shows how custom functions, the ribbon UI, and the task pane code will all run in the same browser runtime.

![Custom functions running in the same browser runtime as the task pane in Excel](../images/custom-functions-in-browser-runtime.png)

## Differences when running cusotm functions in the browser runtime

When you configure your Excel add-in project to run custom functions in the browser runtime, there are a few differences from using the custom function runtime.

### Storage

You no longer need to use the **Storage** API to share data between the task pane, custom functions or ribbon UI. You can put global variables in the **window** object, or use your own preferred state management approach.

### Authentication

When you receive tokens as part of authentication, you don't need to use the **Storage** API to share them between the task pane, custom functions and ribbon UI. You can use your own preferred storage technique to share them.

### Dialog API

You no longer need to use the **OfficeRuntime.Dialog** API to display a dialog from a custom function. You can use the same **Dialog** API as the task pane.

### Debugging

You can't use VS Code to debug custom functions in Excel on Windows at this time. You'll need to use developer tools. For more information, see [Debug add-ins using developer tools on Windows 10](../testing/debug-add-ins-using-f12-developer-tools-on-windows-10.md).

## Get Started

To configure your Excel add-in project to run custom functions in the browser runtime, see [Tutorial: Share data and events between Excel custom functions and the task pane (preview)](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md).

## Give us feedback

We'd love to hear your feedback on this feature. If you find any bugs, issues, or have requests on this feature, please let us know by creating a GitHub issue in the [office-js repo](https://github.com/OfficeDev/office-js).

## See also

List of related articles for shared runtime
- [Tutorial: Share data and events between Excel custom functions and the task pane (preview)](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [Call Excel APIs from your custom function (preview)](call-excel-apis-from-custom-function.md)
- [Run code when the task pane is closed or not visible (preview)](run-code-on-document-open.md)

