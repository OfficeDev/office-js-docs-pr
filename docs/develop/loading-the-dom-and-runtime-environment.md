---
title: Loading the DOM and runtime environment
description: Load the DOM and Office Add-ins runtime environment.
ms.date: 05/20/2023
ms.localizationpriority: medium
---


# Load the DOM and runtime environment

Before running its own custom logic, an add-in must ensure that both the DOM and the Office Add-ins [runtime](../testing/runtimes.md) environment are loaded.

## Startup of a content or task pane add-in

The following figure shows the flow of events involved in starting a content or task pane add-in in Excel, PowerPoint, Project, or Word.

![Flow of events when starting a content or task pane add-in.](../images/office15-app-sdk-loading-dom-agave-runtime.png)

The following events occur when a content or task pane add-in starts.

1. The user opens a document that already contains an add-in or inserts an add-in in the document.

2. The Office client application reads the add-in's manifest from Microsoft Marketplace, an app catalog on SharePoint, or the shared folder catalog it originates from.

3. The Office client application opens the add-in's HTML page in a webview control.

    The next two steps, steps 4 and 5, occur asynchronously and in parallel. For this reason, your add-in's code must make sure that both the DOM and the add-in runtime environment have finished loading before proceeding.

4. The webview control loads the DOM and HTML body, and calls the event handler for the `window.onload` event.

5. The Office client application loads the runtime environment, which downloads and caches the Office JavaScript API library files from the content distribution network (CDN) server, and then calls the add-in's event handler for the [initialize](/javascript/api/office#Office_initialize_reason_) event of the [Office](/javascript/api/office) object, if a handler has been assigned to it. At this time it also checks to see if any callbacks (or chained `then()` method) have been passed (or chained) to the `Office.onReady` handler. For more information about the distinction between `Office.initialize` and `Office.onReady`, see [Initialize your add-in](initialize-add-in.md).

6. When the DOM and HTML body finish loading and the add-in finishes initializing, the main function of the add-in can proceed.

## Startup of an Outlook add-in

The following figure shows the flow of events involved in starting an Outlook add-in running on the desktop, tablet, or smartphone.

![Flow of events when starting Outlook add-in.](../images/outlook15-loading-dom-agave-runtime.png)

The following events occur when an Outlook add-in starts.

1. When Outlook starts, Outlook reads the manifests for Outlook add-ins that have been installed for the user's email account.

2. The user selects an item in Outlook.

3. If the selected item satisfies the activation conditions of an Outlook add-in, Outlook activates the add-in and makes its button visible in the UI.

4. If the user clicks the button to start the Outlook add-in, Outlook opens the HTML page in a webview control. The next two steps, steps 5 and 6, occur in parallel.

5. The webview control loads the DOM and HTML body, and calls the event handler for the `onload` event.

6. Outlook loads the runtime environment, which downloads and caches the JavaScript API for JavaScript library files from the content distribution network (CDN) server, and then calls the event handler for the [initialize](/javascript/api/office#Office_initialize_reason_) event of the [Office](/javascript/api/office) object of the add-in, if a handler has been assigned to it. At this time it also checks to see if any callbacks (or chained `then()` methods) have been passed (or chained) to the `Office.onReady` handler. For more information about the distinction between `Office.initialize` and `Office.onReady`, see [Initialize your add-in](initialize-add-in.md).

7. When the DOM and HTML body finish loading and the add-in finishes initializing, the main function of the add-in can proceed.

## See also

- [Understanding the Office JavaScript API](understanding-the-javascript-api-for-office.md)
- [Initialize your Office Add-in](initialize-add-in.md)
- [Runtimes in Office Add-ins](../testing/runtimes.md)
