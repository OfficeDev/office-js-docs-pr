---
ms.date: 08/13/2020
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

When you run custom functions in a shared runtime with the task pane, your add-in will run in a Microsoft Internet Explorer 11 browser instance, as explained in [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md). Additionally, any buttons that your Excel add-in displays on the ribbon will run in the same shared runtime. The following image shows how custom functions, the ribbon UI, and the task pane code will all run in the same JavaScript runtime.

![Custom functions running in a shared runtime with ribbon buttons and the task pane in Excel](../images/custom-functions-in-browser-runtime.png)

## Set up a shared runtime

See the [configuring a shared runtime article](./configure-your-add-in-to-use-a-shared-runtime.md) to learn how to set up your custom functions to use a shared runtime.

### Debugging

When using a shared runtime, you can't use Visual Studio Code to debug custom functions in Excel on Windows at this time. You'll need to use developer tools instead. For more information, see [Debug add-ins using developer tools on Windows 10](../testing/debug-add-ins-using-f12-developer-tools-on-windows-10.md).

## Give us feedback

We'd love to hear your feedback on this feature. If you find any bugs, issues, or have requests on this feature, please let us know by creating a GitHub issue in the [office-js repo](https://github.com/OfficeDev/office-js).

## See also

- [Tutorial: Share data and events between Excel custom functions and the task pane](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [Call Excel APIs from your custom function](call-excel-apis-from-custom-function.md)
