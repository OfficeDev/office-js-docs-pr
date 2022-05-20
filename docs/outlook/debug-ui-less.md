---
title: Debug your UI-less Outlook add-in
description: Learn how to debug your UI-less Outlook add-in.
ms.topic: article
ms.date: 05/19/2022
ms.localizationpriority: medium
---

# Debug your UI-less Outlook add-in

This article describes how to use the Office Add-in Debugger Extension in Visual Studio Code to debug [UI-less Outlook add-ins](add-in-commands-for-outlook.md#executing-a-javascript-function). UI-less add-in actions are initiated through an add-in command button in the ribbon. For more information about add-in commands, see [Add-in commands for Outlook](add-in-commands-for-outlook.md).

This article assumes that you already have an add-in project that you'd like to debug. To create a UI-less add-in to practice debugging, follow the steps in [Tutorial: Build a message compose Outlook add-in](../tutorials/outlook-tutorial.md).

## Mark your add-in for debugging

If you used the [Yeoman generator for Office Add-ins](../develop/yeoman-generator-overview.md) to create your add-in project, skip to the [Configure and run the debugger](#configure-and-run-the-debugger) section below. When you run `npm start` to build your add-in and start the local server, the command also sets the `UseDirectDebugger` value of the `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]` registry key to mark your add-in for debugging.

Otherwise, if you used another tool to create your add-in, perform the following steps.

1. Navigate to the `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]` registry key. Replace `[Add-in ID]` with the **Id** from your add-in's manifest.

    [!include[Developer registry key](../includes/developer-registry-key.md)]

1. Set the key's `UseDirectDebugger` value to `1`.

## Configure and run the debugger

Now that you've enabled debugging on your add-in, you're ready to configure and run the debugger. For instructions on how to do this, select one of the following options that applies to your runtime.

- If your add-in runs in the WebView runtime, refer to [Microsoft Office Add-in Debugger Extension for Visual Studio Code](../testing/debug-with-vs-extension.md).

- If your add-in runs in the Microsoft Edge Chromium WebView2 runtime, refer to [Debug add-ins on Windows using Visual Studio Code and Microsoft Edge WebView2 (Chromium-based)](../testing/debug-desktop-using-edge-chromium.md).

## See also

- [Add-in commands for Outlook](add-in-commands-for-outlook.md)
- [Overview of debugging Office Add-ins](../testing/debug-add-ins-overview.md)
- [Debug your event-based Outlook add-in](debug-autolaunch.md)
