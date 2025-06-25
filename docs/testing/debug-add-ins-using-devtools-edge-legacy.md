---
title: Debug add-ins using developer tools for Microsoft Edge Legacy
description: Debug add-ins using the developer tools in Microsoft Edge Legacy.
ms.date: 06/17/2025
ms.localizationpriority: medium
---

# Debug add-ins using developer tools in Microsoft Edge Legacy

This article shows how to debug the client-side code (JavaScript or TypeScript) of your add-in when your computer is using a combination of Windows and Office versions that use the original Edge webview control, EdgeHTML.

To determine which browser or webview you're using, see [Browsers and webview controls used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).

> [!TIP]
> [!INCLUDE[Identify the webview through the add-in UI](../includes/identify-webview-in-ui.md)]

> [!NOTE]
> To install a version of Office that uses the Edge legacy webview or to force your current version of Office to use Edge Legacy, see [Switch to the Edge Legacy webview](#switch-to-the-edge-legacy-webview).

## Debug a task pane add-in using Microsoft Edge DevTools Preview

1. Install the [Microsoft Edge DevTools Preview](https://apps.microsoft.com/detail/9mzbfrmz0mnj). (The word "Preview" is in the name for historical reasons. There isn't a more recent version.)

   > [!NOTE]
   > If your add-in has an [add-in command](../design/add-in-commands.md) that executes a function, the function runs in a hidden browser runtime process that the Microsoft Edge DevTools cannot detect or attach to, so the technique described in this article cannot be used to debug code in the function.

1. [Sideload](test-debug-non-local-server.md) and run the add-in.
1. Run the Microsoft Edge DevTools.
1. In the tools, open the **Local** tab. Your add-in is listed by its name. (Only processes that are running in EdgeHTML appear on the tab. The tool can't attach to processes that are running in other browsers or webviews, including Microsoft Edge (WebView2) and Internet Explorer (Trident).)

   :::image type="content" source="../images/edge-devtools-with-add-in-process.png" alt-text="Edge DevTools showing a process named legacy-edge-debugging.":::

1. Select the add-in name to open it in the tools.
1. Open the **Debugger** tab.
1. Open the file that you want to debug with the following steps.

   1. On the debugger task bar, select **Show find in files**. This action opens a search window.
   1. Enter a line of code from the file you want to debug in the search box. It should be something that's not likely to be in any other file.
   1. Select the refresh button.
   1. In the search results, select the line to open the code file in the pane above the search results.

   :::image type="content" source="../images/open-file-in-edge-devtools.png" alt-text="Edge DevTools debugging tab with 4 parts labelled A through D.":::

1. To set a breakpoint, select the line in the code file. The breakpoint is registered in the **Call stack** (bottom right) pane. There may also be a red dot by the line in the code file, but this doesn't appear reliably.
1. Execute functions in the add-in as needed to trigger the breakpoint.

> [!TIP]
> For more information about using the tools, see [Microsoft Edge (EdgeHTML) Developer Tools](/archive/microsoft-edge/legacy/developer/devtools-guide/).

## Debug a dialog in an add-in

If your add-in uses the Office Dialog API, the dialog runs in a separate process from the task pane (if any) and the tools must attach to that process. Follow these steps.

1. Run the add-in and the tools.
1. Open the dialog and then select the **Refresh** button in the tools. The dialog process is shown. Its name comes from the `<title>` element in the HTML file that is open in the dialog.
1. Select the process to open it and debug just as described in the section [Debug a task pane add-in using Microsoft Edge DevTools Preview](#debug-a-task-pane-add-in-using-microsoft-edge-devtools-preview).

   :::image type="content" source="../images/edge-devtools-with-add-in-and-dialog-processes.png" alt-text="Edge DevTools showing a process named My Dialog.":::

## Switch to the Edge Legacy webview

There are two ways to switch the Edge Legacy webview. You can run a simple command in a command prompt, or you can install a version of Office that uses Edge Legacy by default. We recommend the first method. But you should use the second in the following scenarios.

- Your project was developed with Visual Studio and IIS. It isn't Node.js based.
- You want to be absolutely robust in your testing.
- If for any reason the command line tool doesn't work.

### Switch via the command line

[!INCLUDE [Steps to switch browsers with the command line tool](../includes/use-legacy-edge-or-ie.md)]

### Install a version of Office that uses Edge Legacy

[!INCLUDE [Steps to install Office that uses Edge Legacy or Internet Explorer](../includes/install-office-that-uses-legacy-edge-or-ie.md)]
