---
title: Best practices and rules for the Office dialog API
description: 'Provides rules and best practices for the Office dialog API, such as best practices for a single-page application (SPA)'
ms.date: 01/29/2020
localization_priority: Normal
---

# Best practices and rules for the Office dialog API

This article provides rules, gotchas, and best practices for the Office dialog API, including best practices for designing the UI of a dialog and using the API with in a single-page application (SPA)

> [!NOTE]
> This article presupposes that you are familiar with the basics of using the Office dialog API as described in [Use the Office dialog API in your Office Add-ins](dialog-api-in-office-add-ins.md).
> 
> See also [Handling errors and events with the Office dialog box](dialog-handle-errors-events.md).

## Rules and gotchas

- The dialog box can only navigate to HTTPS URLs, not HTTP.
- The URL passed to the [displayDialogAsync](/javascript/api/office/office.ui) method must be in the exact same domain as the add-in itself. It cannot be a subdomain. But the page that is passed to it can redirect to a page in another domain.
- A host window, which can be a task pane or the UI-less [function file](../reference/manifest/functionfile.md) of an add-in command, can have only one dialog box open at a time.
- Only two Office APIs can be called in the dialog box:
  - The [messageParent](/javascript/api/office/office.ui#messageparent-message-) function.
  - `Office.context.requirements.isSetSupported` (For more information, see [Specify Office applications and API requirements](specify-office-hosts-and-api-requirements.md).)
- The [messageParent](/javascript/api/office/office.ui#messageparent-message-) function can only be called from a page in the exact same domain as the add-in itself.

## Best practices

### Avoid overusing dialog boxes

Because overlapping UI elements are discouraged, avoid opening a dialog box from a task pane unless your scenario requires it. When you consider how to use the surface area of a task pane, note that task panes can be tabbed. For an example, see the [Excel Add-in JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) sample.

### Designing a dialog box UI

For best practices in dialog box design, see [Dialog boxes in Office Add-ins](../design/dialog-boxes.md).

### Handling pop-up blockers with Office on the web

Attempting to display a dialog box while using Office on the web may cause the browser's pop-up blocker to block the dialog box. Office on the web has a feature that enables your add-in's dialog boxes to be an exception to the browser's pop-up blocker. When your code calls the `displayDialogAsync` method, then Office on the web will open a prompt similar to the following.

![The prompt that an add-in can generate to avoid in-browser pop-up blockers.](../images/dialog-prompt-before-open.png)

If the user chooses **Allow**, the Office dialog box opens. If the user chooses **Ignore**, the prompt closes and the Office dialog box does not open. Instead, the `displayDialogAsync` method returns error 12009. Your code should catch this error and either provide an alternate experience that does not require a dialog, or display a message to the user advising that the add-in requires them to allow the dialog. (For more about 12009, see [Errors from displayDialogAsync](dialog-handle-errors-events.md#errors-from-displaydialogasync).)

If, for any reason, you want to turn off this feature, then your code must opt out. It makes this request with the [DialogOptions](/javascript/api/office/office.dialogoptions) object that is passed to the `displayDialogAsync` method. Specifically, the object should include `promptBeforeOpen: false`. When this option is set to false, Office on the web will not prompt the user to allow the add-in open a dialog, and the Office dialog will not open.

### Do not use the \_host\_info value

Office automatically adds a query parameter called `_host_info` to the URL that is passed to `displayDialogAsync`. It is appended after your custom query parameters, if any. It is not appended to any subsequent URLs that the dialog box navigates to. Microsoft may change the content of this value, or remove it entirely, so your code should not read it. The same value is added to the dialog box's session storage. Again, *your code should neither read nor write to this value*.

### Best practices for using the Office dialog API in an SPA

If your add-in uses client-side routing, as single-page applications (SPAs) typically do, you have the option to pass the URL of a route to the [displayDialogAsync](/javascript/api/office/office.ui) method instead of the URL of a separate HTML page. *We recommend against doing so for the reasons given below.*

> [!NOTE]
> This article is not relevant to *server-side* routing, such as in an Express-based web application.

#### Problems with SPAs and the Office dialog API

The Office dialog box is in a new window with its own instance of the JavaScript engine, and hence it's own complete execution context. If you pass a route, your base page and all its initialization and bootstrapping code run again in this new context, and any variables are set to their initial values in the dialog box. So this technique downloads and launches a second instance of your application in the  box window, which partially defeats the purpose of an SPA. In addition, code that changes variables in the dialog box window does not change the task pane version of the same variables. Similarly, the dialog box window has its own session storage, which is not accessible from code in the task pane. The dialog box and the host page on which `displayDialogAsync` was called look like two different clients to your server. (For a reminder of what a host page is, see [Open a dialog box from a host page](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page).)

So, if you passed a route to the `displayDialogAsync` method, you wouldn't really have an SPA; you'd have *two instances of the same SPA*. Moreover, much of the code in the task pane instance would never be used in that instance and much of the code in the dialog box instance would never be used in that instance. It would be like having two SPAs in the same bundle.

#### Microsoft recommendations

Instead of passing a client-side route to the `displayDialogAsync` method, we recommend that you do one of the following:

* If the code that you want to run in the dialog box is sufficiently complex, create two different SPAs explicitly; that is, have two SPAs in different folders of the same domain. One SPA runs in the dialog box and the other in the dialog box's host page where `displayDialogAsync` was called. 
* In most scenarios, only simple logic is needed in the dialog box. In such cases, your project will be greatly simplified by hosting a single HTML page, with embedded or referenced JavaScript, in the domain of your SPA. Pass the URL of the page to the `displayDialogAsync` method. While this means that you are deviating from the literal idea of a single-page app; you don't really have a single instance of an SPA when you are using the Office dialog API.
