---
ms.date: 05/13/2020
description: 'Learn how to run custom functions, ribbon buttons, and task pane code in a the same JavaScript runtime to coordinate scenarios across your add-in.'
title: Run your add-in code in a shared JavaScript runtime
localization_priority: Priority
---

# Overview: Run your add-in code in a shared JavaScript runtimes

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

When running Excel on Windows or Mac, your add-in will run code for ribbon buttons, custom functions, and the task pane in separate JavaScript runtime environments. This creates limitations such as not being able to easily share global data, and not being able to access all CORS functionality from a custom function.

However, you can configure your Excel add-in to share code in the same JavaScript runtime (also referred to as a shared runtime). This enables better coordination across your add-in and access to the task pane DOM and CORS from all parts of your add-in.

Configuring a shared runtime enables the following scenarios:

- Your add-in will have a shared DOM that the ribbon, task pane, and custom functions can all access.
- Your custom functions will have full CORS support.
- Your custom functions can call Office.js APIs to read spreadsheet document data.
- Your add-in can run code as soon as the document is opened.
- Your add-in can continue running code after the task pane is closed.

When you run custom functions in a shared runtime with the task pane, it will run in a browser instance on different platforms as explained in [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md). Additionally, any buttons that your Excel add-in displays on the ribbon will run in the same shared runtime. The following image shows how custom functions, the ribbon UI, and the task pane code will all run in the same JavaScript runtime.

![Custom functions running in the shared runtime with ribbon buttons and the task pane in Excel](../images/custom-functions-in-browser-runtime.png)

## Set up shared runtime

See the [shared runtime tutorial](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md) to learn how to set up your custom functions to use shared runtime.

### Debugging

When using a shared runtime, you can't use Visual Studio Code to debug custom functions in Excel on Windows at this time. You'll need to use developer tools. For more information, see [Debug add-ins using developer tools on Windows 10](../testing/debug-add-ins-using-f12-developer-tools-on-windows-10.md).

## See also

List of related articles for shared runtime
- [Tutorial: Share data and events between Excel custom functions and the task pane (preview)](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [Call Excel APIs from your custom function (preview)](call-excel-apis-from-custom-function.md)