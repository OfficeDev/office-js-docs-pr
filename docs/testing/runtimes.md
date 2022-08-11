---
title: Runtimes in Office Add-ins
description: Learn about the JavaScript runtimes that are used by Office Add-ins.
ms.date: 08/10/2022
ms.localizationpriority: medium
---

# Runtimes in Office Add-ins

As an interpreted language, JavaScript must run in a JavaScript engine. As a single-threaded, synchronous language, JavaScript has no inherent capacity for asynchronous execution; but modern JavaScript engines can request asynchronous operations (including networking communication) from the host operating system and receive data from the OS in response. This kind of engine makes JavaScript *effectively* asynchronous. In this article, engines of this sort are called *runtimes*. [Node.js](https://nodejs.org) and modern browsers are examples of such runtimes. 

## Types of runtimes

There are two types of runtimes used by Office Add-ins:

- **JavaScript-only runtime**: A JavaScript engine supplemented with support for [WebSockets](https://developer.mozilla.org/docs/Web/API/WebSockets_API), [CORS (Cross-Origin Resource Sharing)](https://developer.mozilla.org/docs/Web/HTTP/CORS), and [local storage](https://developer.mozilla.org/docs/Web/API/Window/localStorage). 
- **Browser runtime**: Includes all the features of a JavaScript-only runtime and adds a [rendering engine](https://developer.mozilla.org/docs/Glossary/Rendering_engine) that renders HTML.

Details about these types are later in this article at [JavaScript-only runtime](#javascript-only-runtime) and [Browser runtime](#browser-runtime).

The following table shows which type of runtime is used for the various possible features of an add-in. Note the choice of which type of runtime to use is an implementation detail that Microsoft could change at any time. The Office JavaScript Library doesn't assume that the same type of runtime will always be used for a given feature and your add-in architecture shouldn't assume this either.

| Type of runtime | Add-in feature |
|:-----|:-----|
| JavaScript-only | Excel [custom function](../excel/custom-functions-overview.md)</br>(except when the runtime is [shared](#shared-runtime))</br></br>[Outlook event-based (autolaunched) task](../outlook/autolaunch.md)|
| browser | [task pane](../design/task-pane-add-ins.md)</br></br>[dialog](../develop/dialog-api-in-office-add-ins.md)</br></br>[function command](../design/add-in-commands.md#types-of-add-in-commands)</br></br>Excel [custom function](../excel/custom-functions-overview.md)</br>(when the runtime is [shared](#shared-runtime))|

A dialog always runs in its own process. So does an Outlook event-based task, also called autolaunched task. By default, task panes, function commands, and Excel custom functions each run in their own runtime process. However, for some Office host applications, the add-in manifest can be configured so that any two, or all three, of these run in the same runtime. See [Shared runtime](#shared-runtime).

Depending on the host Office application and the features used in the add-in, there may be as many as four runtimes in an add-in, each running in its own process (but not necessarily running simultaneously). The following are examples.

- A PowerPoint or Word add-in that that doesn't share any runtimes, and includes the following features, has three runtimes.

  - A task pane
  - A function command
  - A dialog (A dialog can be launched from either the task pane or the function command.)

- An Excel add-in that doesn't share any runtimes, and includes the following features, has four runtimes.

  - A task pane
  - A function command
  - A custom function
  - A dialog (A dialog can be launched from either the task pane, the function command, or a custom function.)


- An Excel add-in with the same features and is configured to share the same runtime across the task pane, function command, and custom function, has *two* runtimes.
- An Excel add-in with the same features, except that it has no dialog, and is configured to share the same runtime across the task pane, function command, and custom function, has *one* runtime.
- An Outlook add-in that has the following features has four runtimes. (Runtimes cannot be shared in Outlook.)

  - A task pane
  - A function command
  - A event-based (autolaunched) task
  - A dialog (A dialog can be launched from either the task pane or the function command, but not from an event-based task.)

## Share data across runtimes

For Excel, PowerPoint, and Word add-ins, we recommend using a [Shared runtime](#shared-runtime) when any two or more features, except dialogs, need to share data. In Outlook, or scenarios where sharing a runtime isn't feasible, you need alternative methods. The parts of the add-in that are in separate runtime processes don't share global data automatically and are treated by the add-in's web application server as separate sessions, so [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) cannot be used to share data between them. The following are the recommended ways to share data *between unshared runtimes*.

- Pass data between a dialog and its parent task pane, function command, or custom function by using the [Office.ui.messageParent](/javascript/api/office/office.ui#office-office-ui-messageparent-member(1)) and [Dialog.messageChild](/javascript/api/office/office.dialog#office-office-dialog-messagechild-member(1)) methods.
- To share data between a task pane and a function command, store data in [Window.localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage), which is shared across all runtimes that access the same specific [origin](https://developer.mozilla.org/docs/Glossary/Origin). *LocalStorage isn't accessible in a JavaScript-only runtime and, thus, it isn't available in Excel custom functions or Outlook event-based tasks.*

    > [!NOTE]
    > Data in `Window.localStorage` persists between sessions of the add-in and is shared by add-ins with the same origin. Both of these characteristics are often undesirable for an add-in. You can ensure that each session of a given add-in starts fresh by calling the [Window.localStorage.clear](https://developer.mozilla.org/docs/Web/API/Storage/clear) method when the add-in starts. To allow some stored values to persist, but reinitialize other values, you can use [Window.localStorage.setItem](https://developer.mozilla.org/docs/Web/API/Storage/setItem) when the add-in starts for each item that should be reset to an initial value. You can also call [Window.localStorage.removeItem](https://developer.mozilla.org/docs/Web/API/Storage/removeItem) to delete an item entirely.

- To share data between an Excel custom function and any other runtime, use [OfficeRuntime.storage](/javascript/api/office-runtime/officeruntime.storage).
- To share data between an Outlook event-based task and any other runtime, use [OfficeRuntime.storage](/javascript/api/office-runtime/officeruntime.storage).

Other ways to share data include the following:

- Store shared data in an online database that is accessible to all the runtimes.
- Store shared data in a cookie for the add-in's domain.

For more information, see [Persist add-in state and settings](../develop/persisting-add-in-state-and-settings.md) and [Manage state and settings for an Outlook add-in](../outlook/manage-state-and-settings-outlook.md).

## JavaScript-only runtime

The JavaScript-only runtime that is used in Office add-ins is a modification of an open source runtime originally created for [React Native](https://reactnative.dev/). It contains a JavaScript engine supplemented with support for [WebSockets](https://developer.mozilla.org/docs/Web/API/WebSockets_API), [CORS (Cross-Origin Resource Sharing)](https://developer.mozilla.org/docs/Web/HTTP/CORS), and [local storage](https://developer.mozilla.org/docs/Web/API/Window/localStorage). It doesn't have a rendering engine. 

This type of runtime is used in Outlook event-based (autolaunch) tasks and in Excel custom functions except when the custom function is [sharing a runtime](#shared-runtime). 

- When used for an Excel custom function, the runtime starts up when either the worksheet recalculates or the custom function calculates. It doesn't shut down until the workbook is closed.  
- When used in an Outlook event-based task, the runtime starts up when the event occurs. It ends when the first of the following occurs:

  - The event handler calls the `completed` method of its event parameter.
  - 5 minutes has elapsed since the triggering event.
  - The user changes focus from the window where the event was triggered, such as a message compose window.

## Browser runtime

Office add-ins use a different browser type runtime depending on the platform in which Office is running (web, Mac, or Windows), and on the version and build of Windows and Office. For example, if the user is running Office on the web in a FireFox browser, then the Firefox runtime is used. If the user is running Office on Mac, then the Safari runtime is used. If the user is running Office on Windows, then either an Edge or Internet Explorer provides the runtime, depending on the version of Windows and Office. Details can be found in [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).

All of these runtimes include an HTML rendering engine and provide support for [WebSockets](https://developer.mozilla.org/docs/Web/API/WebSockets_API), [CORS (Cross-Origin Resource Sharing)](https://developer.mozilla.org/docs/Web/HTTP/CORS), and [local storage](https://developer.mozilla.org/docs/Web/API/Window/localStorage). 

A browser runtime lifespan varies depending on the feature that it implements.

- When an add-in with a task pane is launched a browser runtime starts. It shuts down when the add-in is closed.
- When a dialog is opened, a browser runtime starts. It shuts down when the dialog is closed.
- When a function command is executed (which happens when a user selects its button or menu item), a browser runtime starts, unless it is a shared runtime that is already running. If it is a shared runtime it shuts down when the add-in is closed. If it is an unshared runtime, it shuts down when the first of the following occurs:
 
  - The function command calls the `completed` method of its event parameter.
  - 5 minutes has elapsed since the triggering event. (If a dialog was opened in the custom function and it is still open when the parent runtime times-out, the dialog runtime stays running until the dialog is closed.)

- When an Excel custom function is using a shared runtime, then a browser type runtime starts when the custom function calculates if the shared runtime has not already started for some other reason. It shuts down when the add-in is closed.

> [!NOTE]
> When a runtime is being [shared](#shared-runtime), it is possible to close the task pane without shutting down the add-in. See [Show or hide the task pane of your Office Add-in](../develop/show-hide-add-in.md) for more information.

### Shared runtime

A "shared runtime" isn't a type of runtime. It refers to a [browser type runtime](#browser-runtime) that is being shared by features of the add-in that would otherwise each have their own runtime. Specifically, you have the option of configuring the add-in's task pane and function commands to share a runtime. In an Excel add-in, you also can configure a custom function to share the runtime of a task pane or function command or both. When you do this, the custom function is running in a browser type runtime, instead of a [JavaScript-only runtime](#javascript-only-runtime) as it otherwise would. See [Configure your add-in to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md) for information about the benefits and limitations of sharing runtimes and instructions for configuring the add-in to use a shared runtime. 

> [!NOTE]
> - You can share runtimes only in Excel, PowerPoint, and Word. 
> - You cannot configure a dialog to share a runtime. Each dialog always has its own.
> - A shared runtime never uses the original Microsoft Edge WebView (EdgeHTML) runtime. If the conditions for using Microsoft Edge with WebView2 (Chromium-based) are met (as specified in [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md)), then that runtime is used. Otherwise, the Internet Explorer 11 runtime is used.