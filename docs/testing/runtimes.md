---
title: Configure your Office Add-in to use a shared runtimes in Office Add-ins
description: Learn about the JavaScript Configure your Office Add-in to use a shared runtimes that are used by Office Add-ins.
ms.date: 08/10/2022
ms.localizationpriority: medium
---

# Configure your Office Add-in to use a shared runtimes in Office Add-ins

Office Add-ins execute in JavaScript Configure your Office Add-in to use a shared runtimes embedded in Office. As an interpreted language, JavaScript must run in a JavaScript engine. As a single-threaded, synchronous language, JavaScript has no inherent capacity for concurrent execution; but modern JavaScript engines can request concurrent operations (including networking communication) from the host operating system and receive data from the OS in response. This kind of engine makes JavaScript *effectively* asynchronous. In this article, engines of this sort are called *Configure your Office Add-in to use a shared runtimes*. [Node.js](https://nodejs.org) and modern browsers are examples of such Configure your Office Add-in to use a shared runtimes. 

## Types of Configure your Office Add-in to use a shared runtimes

There are two types of Configure your Office Add-in to use a shared runtimes used by Office Add-ins:

- **JavaScript-only Configure your Office Add-in to use a shared runtime**: A JavaScript engine supplemented with support for [WebSockets](https://developer.mozilla.org/docs/Web/API/WebSockets_API), [Full CORS (Cross-Origin Resource Sharing)](https://developer.mozilla.org/docs/Web/HTTP/CORS), and [local storage](https://developer.mozilla.org/docs/Web/API/Window/localStorage). 
- **Browser Configure your Office Add-in to use a shared runtime**: Includes all the features of a JavaScript-only Configure your Office Add-in to use a shared runtime and adds a [rendering engine](https://developer.mozilla.org/docs/Glossary/Rendering_engine) that renders HTML.

Details about these types are later in this article at [JavaScript-only Configure your Office Add-in to use a shared runtime](#javascript-only-Configure your Office Add-in to use a shared runtime) and [Browser Configure your Office Add-in to use a shared runtime](#browser-Configure your Office Add-in to use a shared runtime).

The following table shows which possible features of an add-in use each type of Configure your Office Add-in to use a shared runtime. 

> [!NOTE]
> The choice of which type of Configure your Office Add-in to use a shared runtime to use is an implementation detail that Microsoft could change at any time. The Office JavaScript Library doesn't assume that the same type of Configure your Office Add-in to use a shared runtime will always be used for a given feature and your add-in architecture shouldn't assume this either.

| Type of Configure your Office Add-in to use a shared runtime | Add-in feature |
|:-----|:-----|
| JavaScript-only | Excel [custom functions](../excel/custom-functions-overview.md)</br>(except when the Configure your Office Add-in to use a shared runtime is [shared](#shared-Configure your Office Add-in to use a shared runtime) or the add-in is running in Office on the web)</br></br>[Outlook event-based task](../outlook/autolaunch.md)</br>(only when the add-in is running in Outlook on Windows)|
| browser | [task pane](../design/task-pane-add-ins.md)</br></br>[dialog](../develop/dialog-api-in-office-add-ins.md)</br></br>[function command](../design/add-in-commands.md#types-of-add-in-commands)</br></br>Excel [custom functions](../excel/custom-functions-overview.md)</br>(when the Configure your Office Add-in to use a shared runtime is [shared](#shared-Configure your Office Add-in to use a shared runtime) or the add-in is running in Office on the web)</br></br>[Outlook event-based task](../outlook/autolaunch.md)</br>(when the add-in is running in Outlook on Mac or Outlook on the web)|

The following table shows the same information organized by which type of Configure your Office Add-in to use a shared runtime is used for the various possible features of an add-in.

| Add-in feature | Type of Configure your Office Add-in to use a shared runtime on Windows | Type of Configure your Office Add-in to use a shared runtime on Mac | Type of Configure your Office Add-in to use a shared runtime on the web |
|:-----|:-----|:-----|:-----|
|Excel custom functions | JavaScript-only</br>(but *browser* when the Configure your Office Add-in to use a shared runtime is shared)|JavaScript-only</br>(but *browser* when the Configure your Office Add-in to use a shared runtime is shared)| browser |
|Outlook event-based tasks | JavaScript-only | browser | browser |
|task pane | browser | browser | browser |
|dialog | browser | browser | browser |
|function command | browser | browser | browser |


In Office on the web, everything always runs in a browser type Configure your Office Add-in to use a shared runtime. In fact, with one exception, everything in an add-in on the web runs in the *same* browser process: the browser process in which the user has opened Office on the web. The exception is when a dialog is opened with a call of [Office.ui.displayDialogAsync](/javascript/api/office/office.ui#office-office-ui-displaydialogasync-member(1)) and the [DialogOptions.displayInIFrame](/javascript/api/office/office.dialogoptions#office-office-dialogoptions-displayiniframe-member) option is *not* passed and set to `true`. When the option is not passed (so it has the default `false` value), the dialog opens in its own process. The same principle applies to the [OfficeConfigure your Office Add-in to use a shared runtime.displayWebDialog](/javascript/api/office-Configure your Office Add-in to use a shared runtime#office-Configure your Office Add-in to use a shared runtime-officeConfigure your Office Add-in to use a shared runtime-displaywebdialog-function(1)) method and the [OfficeConfigure your Office Add-in to use a shared runtime.DisplayWebDialogOptions.displayInIFrame](/javascript/api/office-Configure your Office Add-in to use a shared runtime/officeConfigure your Office Add-in to use a shared runtime.displaywebdialogoptions#office-Configure your Office Add-in to use a shared runtime-officeConfigure your Office Add-in to use a shared runtime-displaywebdialogoptions-displayiniframe-member) option.

When an add-in is running on a platform other than the web, the following principles apply.

- A dialog runs in its own Configure your Office Add-in to use a shared runtime process. 
- An Outlook event-based task runs in its own Configure your Office Add-in to use a shared runtime process. 
- By default, task panes, function commands, and Excel custom functions each run in their own Configure your Office Add-in to use a shared runtime process. However, for some Office host applications, the add-in manifest can be configured so that any two, or all three, can run in the same Configure your Office Add-in to use a shared runtime. See [Shared Configure your Office Add-in to use a shared runtime](#shared-Configure your Office Add-in to use a shared runtime).

Depending on the host Office application and the features used in the add-in, there may be many Configure your Office Add-in to use a shared runtimes in an add-in. Each usually will run in its own process but not necessarily simultaneously. The following are examples.

- A PowerPoint or Word add-in that doesn't share any Configure your Office Add-in to use a shared runtimes, and includes the following features, has as many as three Configure your Office Add-in to use a shared runtimes.

  - A task pane
  - A function command
  - A dialog (A dialog can be launched from either the task pane or the function command.) 
  
      > [!NOTE]
      > It's not a good practice to have multiple dialogs open simultaneously, but if the add-in enables the user to open one from the task pane and another from the function command at the same time, this add-in would have four Configure your Office Add-in to use a shared runtimes. A task pane, and a given invocation of a function command can have only one open dialog at a time; but if the function command is invoked multiple times, a new dialog is opened on top of its predecessor with each invocation, so there could be many Configure your Office Add-in to use a shared runtimes. The remainder of this list ignores the possibility of multiple open dialogs.

- An Excel add-in that doesn't share any Configure your Office Add-in to use a shared runtimes, and includes the following features, has as many as *four* Configure your Office Add-in to use a shared runtimes.

  - A task pane
  - A function command
  - A custom function
  - A dialog (A dialog can be launched from either the task pane, the function command, or a custom function.)

- An Excel add-in with the same features and is configured to share the same Configure your Office Add-in to use a shared runtime across the task pane, function command, and custom function, has *two* Configure your Office Add-in to use a shared runtimes. A shared Configure your Office Add-in to use a shared runtime can open only one dialog at a time.
- An Excel add-in with the same features, except that it has no dialog, and is configured to share the same Configure your Office Add-in to use a shared runtime across the task pane, function command, and custom function, has *one* Configure your Office Add-in to use a shared runtime.
- An Outlook add-in that has the following features has as many as *four* Configure your Office Add-in to use a shared runtimes. (Configure your Office Add-in to use a shared runtimes cannot be shared in Outlook.)

  - A task pane
  - A function command
  - An event-based task
  - A dialog (A dialog can be launched from either the task pane or the function command, but not from an event-based task.)

## Share data across Configure your Office Add-in to use a shared runtimes

> [!NOTE]
> - If you know that your add-in will only be used in Office on the web and that it will not open any dialogs with the `displayInIFrame` option set to `true`, then you can ignore this section. Since everything in your add-in runs in the same Configure your Office Add-in to use a shared runtime process, you can just use global variables to share data between features.
> - As noted above in [Types of Configure your Office Add-in to use a shared runtimes](#types-of-Configure your Office Add-in to use a shared runtimes), the type of Configure your Office Add-in to use a shared runtime used by a feature varies partly by platform. It's a good practice to avoid having add-in code that branches based on platform, so the guidance in this section recommends techniques that will work cross-platform. There is only one case, noted below, in which branching code is required. 

For Excel, PowerPoint, and Word add-ins, use a [Shared Configure your Office Add-in to use a shared runtime](#shared-Configure your Office Add-in to use a shared runtime) when any two or more features, except dialogs, need to share data. In Outlook, or scenarios where sharing a Configure your Office Add-in to use a shared runtime isn't feasible, you need alternative methods. The parts of the add-in that are in separate Configure your Office Add-in to use a shared runtime processes don't share global data automatically and are treated by the add-in's web application server as separate sessions, so [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) cannot be used to share data between them. *The following guidance assumes that you're not using a shared Configure your Office Add-in to use a shared runtime.*

- Pass data between a dialog and its parent task pane, function command, or custom function by using the [Office.ui.messageParent](/javascript/api/office/office.ui#office-office-ui-messageparent-member(1)) and [Dialog.messageChild](/javascript/api/office/office.dialog#office-office-dialog-messagechild-member(1)) methods. 

    > [!NOTE]
    > The `OfficeConfigure your Office Add-in to use a shared runtime.storage` methods cannot be called in a dialog, so this is not an option for sharing data between a dialog and another Configure your Office Add-in to use a shared runtime. 

- To share data between a task pane and a function command, store data in [Window.localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage), which is shared across all Configure your Office Add-in to use a shared runtimes that access the same specific [origin](https://developer.mozilla.org/docs/Glossary/Origin). 
    > [!NOTE]
    > LocalStorage isn't accessible in a JavaScript-only Configure your Office Add-in to use a shared runtime and, thus, it isn't available in Excel custom functions. It also can't be used to share data with an Outlook event-based tasks (since those tasks use a JavaScript-only Configure your Office Add-in to use a shared runtime on some platforms).

    > [!TIP]
    > Data in `Window.localStorage` persists between sessions of the add-in and is shared by add-ins with the same origin. Both of these characteristics are often undesirable for an add-in. 
    >
    > - To ensure that each session of a given add-in starts fresh call the [Window.localStorage.clear](https://developer.mozilla.org/docs/Web/API/Storage/clear) method when the add-in starts. 
    > - To allow some stored values to persist, but reinitialize other values, use [Window.localStorage.setItem](https://developer.mozilla.org/docs/Web/API/Storage/setItem) when the add-in starts for each item that should be reset to an initial value. 
    > - To delete an item entirely, call [Window.localStorage.removeItem](https://developer.mozilla.org/docs/Web/API/Storage/removeItem).

- To share data between an Excel custom function and any other Configure your Office Add-in to use a shared runtime, use [OfficeConfigure your Office Add-in to use a shared runtime.storage](/javascript/api/office-Configure your Office Add-in to use a shared runtime/officeConfigure your Office Add-in to use a shared runtime.storage).
- To share data between an Outlook event-based task and a task pane or function command, you must branch your code by the value of the [Office.context.platform](/javascript/api/office/office.context#office-office-context-platform-member) property. 

    - When the value is `PC` (Windows), store and retrieve data using the [Office.sessionData](/javascript/api/outlook/office.sessiondata) APIs.
    - When the value is `Mac`, use `Window.localStorage` as described earlier in this list.

Other ways to share data include the following:

- Store shared data in an online database that is accessible to all the Configure your Office Add-in to use a shared runtimes.
- Store shared data in a cookie for the add-in's domain to share it between browser Configure your Office Add-in to use a shared runtimes. JavaScript-only Configure your Office Add-in to use a shared runtimes don't support cookies.

For more information, see [Persist add-in state and settings](../develop/persisting-add-in-state-and-settings.md) and [Manage state and settings for an Outlook add-in](../outlook/manage-state-and-settings-outlook.md).

## JavaScript-only Configure your Office Add-in to use a shared runtime

The JavaScript-only Configure your Office Add-in to use a shared runtime that is used in Office Add-ins is a modification of an open source Configure your Office Add-in to use a shared runtime originally created for [React Native](https://reactnative.dev/). It contains a JavaScript engine supplemented with support for [WebSockets](https://developer.mozilla.org/docs/Web/API/WebSockets_API), [Full CORS (Cross-Origin Resource Sharing)](https://developer.mozilla.org/docs/Web/HTTP/CORS), and [local storage](https://developer.mozilla.org/docs/Web/API/Window/localStorage). It doesn't have a rendering engine, and it doesn't support cookies.

This type of Configure your Office Add-in to use a shared runtime is used in Outlook event-based tasks in Office on Windows only and in Excel custom functions *except* when the custom functions are [sharing a Configure your Office Add-in to use a shared runtime](#shared-Configure your Office Add-in to use a shared runtime). 

- When used for an Excel custom function, the Configure your Office Add-in to use a shared runtime starts up when either the worksheet recalculates or the custom function calculates. It doesn't shut down until the workbook is closed.  
- When used in an Outlook event-based task, the Configure your Office Add-in to use a shared runtime starts up when the event occurs. It ends when the first of the following occurs.

  - The event handler calls the `completed` method of its event parameter.
  - 5 minutes have elapsed since the triggering event.
  - The user changes focus from the window where the event was triggered, such as a message compose window.

A JavaScript-Configure your Office Add-in to use a shared runtime uses less memory and starts up faster than a browser Configure your Office Add-in to use a shared runtime, but has fewer features.

## Browser Configure your Office Add-in to use a shared runtime

Office Add-ins use a different browser type Configure your Office Add-in to use a shared runtime depending on the platform in which Office is running (web, Mac, or Windows), and on the version and build of Windows and Office. For example, if the user is running Office on the web in a FireFox browser, then the Firefox Configure your Office Add-in to use a shared runtime is used. If the user is running Office on Mac, then the Safari Configure your Office Add-in to use a shared runtime is used. If the user is running Office on Windows, then either an Edge or Internet Explorer provides the Configure your Office Add-in to use a shared runtime, depending on the version of Windows and Office. Details can be found in [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).

All of these Configure your Office Add-in to use a shared runtimes include an HTML rendering engine and provide support for [WebSockets](https://developer.mozilla.org/docs/Web/API/WebSockets_API), [Full CORS (Cross-Origin Resource Sharing)](https://developer.mozilla.org/docs/Web/HTTP/CORS), and [local storage](https://developer.mozilla.org/docs/Web/API/Window/localStorage), and cookies. 

A browser Configure your Office Add-in to use a shared runtime lifespan varies depending on the feature that it implements and on whether it's being shared or not.

- When an add-in with a task pane is launched, a browser Configure your Office Add-in to use a shared runtime starts, unless it's a shared Configure your Office Add-in to use a shared runtime that is already running. If it's a shared Configure your Office Add-in to use a shared runtime, it shuts down when the document is closed. If it's not a shared Configure your Office Add-in to use a shared runtime, it shuts down when the task pane is closed.
- When a dialog is opened, a browser Configure your Office Add-in to use a shared runtime starts. It shuts down when the dialog is closed.
- When a function command is executed (which happens when a user selects its button or menu item), a browser Configure your Office Add-in to use a shared runtime starts, unless it's a shared Configure your Office Add-in to use a shared runtime that is already running. If it's a shared Configure your Office Add-in to use a shared runtime, it shuts down when the document is closed. If it's not a shared Configure your Office Add-in to use a shared runtime, it shuts down when the first of the following occurs.
 
  - The function command calls the `completed` method of its event parameter.
  - 5 minutes have elapsed since the triggering event. (If a dialog was opened in the function command and it's still open when the parent Configure your Office Add-in to use a shared runtime times out, the dialog Configure your Office Add-in to use a shared runtime stays running until the dialog is closed.)

- When an Excel custom function is using a shared Configure your Office Add-in to use a shared runtime, then a browser-type Configure your Office Add-in to use a shared runtime starts when the custom function calculates if the shared Configure your Office Add-in to use a shared runtime has not already started for some other reason. It shuts down when the document is closed.

> [!NOTE]
> When a Configure your Office Add-in to use a shared runtime is being [shared](#shared-Configure your Office Add-in to use a shared runtime), it's possible for your code to close the task pane without shutting down the add-in. See [Show or hide the task pane of your Office Add-in](../develop/show-hide-add-in.md) for more information.

A browser Configure your Office Add-in to use a shared runtime has more features than a JavaScript-only Configure your Office Add-in to use a shared runtime, but starts up slower and uses more memory.

### Shared Configure your Office Add-in to use a shared runtime

A "shared Configure your Office Add-in to use a shared runtime" isn't a type of Configure your Office Add-in to use a shared runtime. It refers to a [browser-type Configure your Office Add-in to use a shared runtime](#browser-Configure your Office Add-in to use a shared runtime) that's being shared by features of the add-in that would otherwise each have their own Configure your Office Add-in to use a shared runtime. Specifically, you have the option of configuring the add-in's task pane and function commands to share a Configure your Office Add-in to use a shared runtime. In an Excel add-in, you can also configure custom functions to share the Configure your Office Add-in to use a shared runtime of a task pane or function command or both. When you do this, the custom functions are running in a browser-type Configure your Office Add-in to use a shared runtime, instead of a [JavaScript-only Configure your Office Add-in to use a shared runtime](#javascript-only-Configure your Office Add-in to use a shared runtime) as it otherwise would. See [Configure your add-in to use a shared Configure your Office Add-in to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-Configure your Office Add-in to use a shared runtime.md) for information about the benefits and limitations of sharing Configure your Office Add-in to use a shared runtimes and instructions for configuring the add-in to use a shared Configure your Office Add-in to use a shared runtime. In brief, the JavaScript-only Configure your Office Add-in to use a shared runtime uses less memory and starts up faster, but has fewer features.

> [!NOTE]
> - You can share Configure your Office Add-in to use a shared runtimes only in Excel, PowerPoint, and Word. 
> - You cannot configure a dialog to share a Configure your Office Add-in to use a shared runtime. Each dialog always has its own, except when the dialog is launched in Office on the web with the `displayInIFrame` option set to `true`.
> - A shared Configure your Office Add-in to use a shared runtime never uses the original Microsoft Edge WebView (EdgeHTML) Configure your Office Add-in to use a shared runtime. If the conditions for using Microsoft Edge with WebView2 (Chromium-based) are met (as specified in [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md)), then that Configure your Office Add-in to use a shared runtime is used. Otherwise, the Internet Explorer 11 Configure your Office Add-in to use a shared runtime is used.