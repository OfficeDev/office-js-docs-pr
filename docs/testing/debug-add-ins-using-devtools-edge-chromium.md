---
title: Debug add-ins using developer tools for Microsoft Edge WebView2
description: Debug add-ins using the developer tools in Microsoft Edge WebView2 (Chromium-based).
ms.date: 09/18/2025
ms.localizationpriority: medium
---

# Debug add-ins using developer tools in Microsoft Edge (Chromium-based)

This article shows how to debug the client-side code (JavaScript or TypeScript) of your add-in when the following conditions are met.

- You can't, or don't wish to, debug using tools built into your IDE; or you are encountering a problem that only occurs when the add-in is run outside the IDE.
- Your computer is using a combination of Windows and Office versions that use the Edge (Chromium-based) webview control, WebView2.

> [!TIP]
> For information about debugging with Edge WebView2 (Chromium-based) inside Visual Studio Code, see [Debug add-ins on Windows using Visual Studio Code and Microsoft Edge WebView2 (Chromium-based)](debug-desktop-using-edge-chromium.md).

To determine which webview you're using, see [Browsers and webview controls used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).

> [!TIP]
> [!INCLUDE[Identify the webview through the add-in UI](../includes/identify-webview-in-ui.md)]

## Debug a task pane add-in using Microsoft Edge (Chromium-based) developer tools

> [!NOTE]
> If your add-in has an [add-in command](../design/add-in-commands.md) that executes a function, the function runs in a hidden browser runtime process that the Microsoft Edge (Chromium-based) developer tools can't be launched from, so the technique described in this article can't be used to debug code in the function.

1. [Sideload](test-debug-non-local-server.md) and run the add-in.

    > [!NOTE]
    > To sideload an add-in in Outlook, see [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md).

1. Run the Microsoft Edge (Chromium-based) developer tools by one of these methods:

   - Be sure the add-in's task pane has focus and press <kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>I</kbd>.
   - Right-click (or select and hold) the task pane to open the context menu and select **Inspect**, or open the [personality menu](../design/task-pane-add-ins.md#personality-menu) and select **Attach Debugger**. (The personality menu isn't supported in Outlook.)

   > [!NOTE]
   > The new Outlook on Windows desktop client doesn't support the context menu or the keyboard shortcut to access the Microsoft Edge developer tools. Instead, you must run `olk.exe --devtools` from a command prompt. For more information, see the "Debug your add-in" section of [Develop Outlook add-ins for the new Outlook on Windows](../outlook/one-outlook.md#debug-your-add-in).

1. Open the **Sources** tab.
1. Open the file that you want to debug with the following steps.

   1. On the far right of the tool's top menu bar, select the **...** button and then select **Search**.
   1. Enter a line of code from the file you want to debug in the search box. It should be something that's not likely to be in any other file.
   1. Select the refresh button.
   1. In the search results, select the line to open the code file in the pane above the search results.

   :::image type="content" source="../images/open-file-in-edge-chromium-devtools.png" alt-text="Edge Chromium developer tools source tab with 4 parts labelled A through D.":::

1. To set a breakpoint, select the line number of the line in the code file. A red dot appears by the line in the code file. In the debugger window to the right, the breakpoint is registered in the **Breakpoints** drop down.
1. Execute functions in the add-in as needed to trigger the breakpoint.

> [!TIP]
> For more information about using the tools, see [Microsoft Edge Developer Tools overview](/microsoft-edge/devtools-guide-chromium/).

## Debug a dialog in an add-in

If your add-in uses the Office Dialog API, the dialog runs in a separate process from the task pane (if any) and the tool must be started from that separate process. Follow these steps.

1. Run the add-in.
1. Open the dialog and be sure it has focus.
1. Open the Microsoft Edge (Chromium-based) developer tools by one of these methods:

   - Press <kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>I</kbd> or <kbd>F12</kbd>.
   - Right-click (or select and hold) the dialog to open the context menu and select **Inspect**.

1. Use the tool the same as you would for code in a task pane. See [Debug a task pane add-in using Microsoft Edge (Chromium-based) developer tools](#debug-a-task-pane-add-in-using-microsoft-edge-chromium-based-developer-tools) earlier in this article.

## Automatically open the Microsoft Edge (Chromium-based) developer tools to debug initialization

[!INCLUDE[Automatically open the Microsoft Edge (Chromium-based) developer tools to debug initialization](../includes/auto-open-webview2-dev-tools.md)]


