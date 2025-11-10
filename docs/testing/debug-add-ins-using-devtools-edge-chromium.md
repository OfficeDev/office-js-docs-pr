---
title: Debug add-ins using developer tools for Microsoft Edge WebView2
description: Debug add-ins using the developer tools in Microsoft Edge WebView2.
ms.date: 09/18/2025
ms.localizationpriority: medium
---

# Debug add-ins using developer tools in Microsoft Edge

This article shows how to debug the client-side code (JavaScript or TypeScript) of your add-in outside of your current IDE.

## Debug a task pane add-in using Microsoft Edge developer tools

> [!NOTE]
> If your add-in has an [add-in command](../design/add-in-commands.md) that executes a function, the function runs in a hidden browser runtime process that the Microsoft Edge developer tools can't be launched from, so the technique described in this article can't be used to debug code in the function.

1. [Sideload](test-debug-non-local-server.md) and run the add-in.

    > [!NOTE]
    > To sideload an add-in in Outlook, see [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md).

1. Run the Microsoft Edge developer tools by one of these methods:

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

   :::image type="content" source="../images/open-file-in-edge-chromium-devtools.png" alt-text="Edge developer tools source tab with 4 parts labelled A through D.":::

1. To set a breakpoint, select the line number of the line in the code file. A red dot appears by the line in the code file. In the debugger window to the right, the breakpoint is registered in the **Breakpoints** drop down.
1. Execute functions in the add-in as needed to trigger the breakpoint.

> [!TIP]
> For more information about using the tools, see [Microsoft Edge Developer Tools overview](/microsoft-edge/devtools-guide-chromium/).

## Debug a dialog in an add-in

If your add-in uses the Office Dialog API, the dialog runs in a separate process from the task pane (if any) and the tool must be started from that separate process. Follow these steps.

1. Run the add-in.
1. Open the dialog and be sure it has focus.
1. Open the Microsoft Edge developer tools by one of these methods:

   - Press <kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>I</kbd> or <kbd>F12</kbd>.
   - Right-click (or select and hold) the dialog to open the context menu and select **Inspect**.

1. Use the tool the same as you would for code in a task pane. See [Debug a task pane add-in using Microsoft Edge developer tools](#debug-a-task-pane-add-in-using-microsoft-edge-developer-tools) earlier in this article.

## Automatically open the Microsoft Edge developer tools to debug initialization

[!INCLUDE[Automatically open the Microsoft Edge developer tools to debug initialization](../includes/auto-open-webview2-dev-tools.md)]
