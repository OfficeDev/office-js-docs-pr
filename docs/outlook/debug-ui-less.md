---
title: Debug function commands in Outlook add-ins
description: Learn how to debug function commands in Outlook add-ins.
ms.date: 07/11/2022
ms.topic: how-to
ms.localizationpriority: medium
---

# Debug function commands in Outlook add-ins

> [!NOTE]
> The technique in this article can be used only on a Windows development computer. If you're developing on a Mac, see [Debug function commands](../testing/debug-function-command.md).

This article describes how to use the Office Add-in Debugger Extension in Visual Studio Code to debug [function commands](../design/add-in-commands.md#types-of-add-in-commands). Function commands are initiated through an add-in command button on the ribbon. For more information about add-in commands, see [Add-in commands](../design/add-in-commands.md).

This article assumes that you already have an add-in project that you'd like to debug. To create an add-in with a function command to practice debugging, follow the steps in [Tutorial: Build a message compose Outlook add-in](../tutorials/outlook-tutorial.md).

## Mark your add-in for debugging

If you used the [Yeoman generator for Office Add-ins](../develop/yeoman-generator-overview.md) to create your add-in project, skip to the [Configure and run the debugger](#configure-and-run-the-debugger) section later in this article. When you run `npm start` to build your add-in and start the local server, the command also sets the `UseDirectDebugger` value of the `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]` registry key to mark your add-in for debugging.

Otherwise, if you used another tool to create your add-in, perform the following steps.

1. Navigate to the `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]` registry key. Replace `[Add-in ID]` with the **\<Id\>** from your add-in's manifest.

    [!include[Developer registry key](../includes/developer-registry-key.md)]

1. Set the key's `UseDirectDebugger` value to `1`.

## Configure and run the debugger

Now that you've enabled debugging on your add-in, you're ready to configure and run the debugger. For instructions on how to do this, select one of the following options that applies to your webview control. For information about how to determine what webview control is used on your development computer, see [Browsers and webview controls used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).

- If your add-in runs in the embedded webview control from Edge Legacy (EdgeHTML), see [Debug add-ins using developer tools in Microsoft Edge Legacy](../testing/debug-add-ins-using-devtools-edge-legacy.md).

- If your add-in runs in the embedded webview control from Microsoft Edge Chromium (WebView2), see [Debug add-ins on Windows using Visual Studio Code and Microsoft Edge WebView2 (Chromium-based)](../testing/debug-desktop-using-edge-chromium.md).

## See also

- [Add-in commands](../design/add-in-commands.md)
- [Overview of debugging Office Add-ins](../testing/debug-add-ins-overview.md)
- [Debug event-based and spam-reporting add-ins](../testing/debug-autolaunch.md)
