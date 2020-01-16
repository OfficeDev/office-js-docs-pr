---
title: Best practices and Rules for the Office Dialog
description: 'Provides rules and best practices for the Office dialog, such as best practices for a single-page application (SPA)'
ms.date: 01/16/2020
localization_priority: Normal
---

# Best practices and Rules for the Office Dialog

> [!NOTE]
> This article presupposes that you are familiar with the basics of using the Office Dialog as described in [Use the Dialog API in your Office Add-ins](dialog-api-in-office-add-ins.md).

## Rules and gotchas

- The dialog can only navigate to HTTP**S** URLs, not HTTP.
- The URL passed to the [displayDialogAsync](/javascript/api/office/office.ui) method must be in the exact same domain as the add-in itself. It cannot be a subdomain. But the page that is passed to it can redirect to a page in another domain.
- A host window, which can be a task pane or the UI-less [function file](/office/dev/add-ins/reference/manifest/functionfile) of an add-in command, can have only one dialog open at a time.
- Only two Office APIs that can be called in the dialog box: the [messageParent](/javascript/api/office/office.ui#messageparent-message-) function  and `Office.context.requirements.isSetSupported`. For information about the second, see [Specify Office hosts and API requirements](specify-office-hosts-and-api-requirements.md).
- The [messageParent](/javascript/api/office/office.ui#messageparent-message-) function can only be called from a page in the exact same domain as the add-in itself.

## Best practices

### Avoid overusing dialogs

Because overlapping UI elements are discouraged, avoid opening a dialog from a task pane unless your scenario requires it. When you consider how to use the surface area of a task pane, note that task panes can be tabbed. For an example, see the [Excel Add-in JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) sample.

### Designing a dialog UI

For best practices in dialog design, see [Dialog boxes in Office Add-ins](/design/dialog-boxes.md).

### Handling pop-up blockers with Office on the web

Attempting to display a dialog while using Office on the web may cause the browser's pop-up blocker to block the dialog. The browser's pop-up blocker can be circumvented if the user of your add-in first agrees to a prompt from the add-in. `displayDialogAsync`'s [DialogOptions](/javascript/api/office/office.dialogoptions) has the `promptBeforeOpen` property to trigger such a pop-up. `promptBeforeOpen` is a boolean value with the following meaning:

 - `true` - The framework displays a pop-up to trigger the navigation and avoid the browser's pop-up blocker. 
 - `false` - The dialog will not be shown and the developer must handle pop-ups (by providing a user interface artifact to trigger the navigation). 
 
The pop-up looks similar to that in the following screenshot:

![The prompt an add-in's dialog can generate to avoid in-browser pop-up blockers.](../images/dialog-prompt-before-open.png)

### Do not use the _host_info value

Office automatically adds a query parameter called `_host_info` to the URL that is passed to `displayDialogAsync`. (It is appended after your custom query parameters, if any. It is not appended to any subsequent URLs that the dialog box navigates to.) Microsoft may change the content of this value, or remove it entirely, in the future, so your code should not read it. The same value is added to the dialog box's session storage. Again, *your code should neither read nor write to this value*.

### Best practices for using the Office Dialog in an SPA

If your add-in uses client-side routing, as single-page applications (SPAs) typically do, you have the option to pass the URL of a route to the [displayDialogAsync](/javascript/api/office/office.ui) method instead of the URL of a complete and separate HTML page. *We recommend against doing so for the reasons given below.*

> [!NOTE]
> This article is not relevant to *server-side* routing, such as in an Express-based web application.

#### Problems with SPAs and the Office Dialog

The dialog box is in a new window with its own instance of the JavaScript engine, and hence it's own complete execution context. If you pass a route, your base page and all its initialization and bootstrapping code run again in this new context, and any variables are set to their initial values in the dialog window. So this technique downloads and launches a second instance of your application in the dialog window, which partially defeats the purpose of an SPA. In addition, code that changes variables in the dialog window does not change the task pane version of the same variables. Similarly, the dialog window has its own session storage, which is not accessible from code in the task pane. The dialog and the host page on which `displayDialogAsync` was called look like two different clients to your server.

So, if you passed a route to the `displayDialogAsync` method, you wouldn't really have an SPA; you'd have *two instances of the same SPA*. Moreover, much of the code in the task pane instance would never be used in that instance and much of the code in the dialog instance would never be used in that instance. It would be like having two SPAs in the same bundle.

#### Microsoft recommendations

Instead of passing a client-side route to the `displayDialogAsync` method, we recommend that you do one of the following:

* If the code that you want to run in the dialog is sufficiently complex, create two different SPAs explicitly; that is, have two SPAs in different folders of the same domain. One SPA runs in the dialog and the other in the dialog's host page where `displayDialogAsync` was called. 
* In most scenarios, only simple logic is needed in the dialog. In such cases, your project will be greatly simplified by simply hosting a simple HTML page, with embedded or referenced JavaScript, in the domain of your SPA. Pass the URL of the page to the `displayDialogAsync` method. This might mean that you are deviating from the literal idea of a single-page app; but as noted above you don't really have a single instance of an SPA anyway when you are using the dialog.
