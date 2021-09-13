---
title: Debug add-ins on Windows using Microsoft Edge WebView2 (Chromium-based)
description: 'Learn how to debug Office Add-ins that use Microsoft Edge WebView2 (Chromium-based) by using the Debugger for Microsoft Edge extension in VS Code.'
ms.date: 08/18/2021
ms.localizationpriority: high
---
# Debug add-ins on Windows using Edge Chromium WebView2

Office Add-ins running on Windows can use the Debugger for Microsoft Edge extension in VS Code to debug against the Edge Chromium WebView2 runtime.

## Prerequisites

- [Visual Studio Code](https://code.visualstudio.com/) (must be run as an administrator)
- [Node.js (version 10+)](https://nodejs.org/)
- Windows 10
- A combination of platform and Office application that supports Microsoft Edge with WebView2 (Chromium-based) as explained in [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md). If your version of Microsoft 365 is earlier than 2101, you will need to install WebView2. Use the instructions for installing it at [Microsoft Edge WebView2 / Embed web content ... with Microsoft Edge WebView2](https://developer.microsoft.com/microsoft-edge/webview2/).

## Install and use the debugger

1. Create a project using the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office). You can use any one of our quick start guides, such as the [Outlook add-in quickstart](../quickstarts/outlook-quickstart.md), in order to do this.

    > [!TIP]
    > If you aren't using a Yeoman generator based add-in, you may be prompted to adjust a registry key. While in the root folder of your project, run the following in the command line:  `office-add-in-debugging start <your manifest path>`

1. Open your project in VS Code. Within VS Code, select **Ctrl+Shift+X** to open the Extensions bar. Search for the "Debugger for Microsoft Edge" extension and install it.

1. Next, choose  **View > Run** or enter **Ctrl+Shift+D** to switch to debug view.

1. From the **RUN AND DEBUG** options, choose the Edge Chromium option for your host application, such as **Excel Desktop (Edge Chromium)**. Select **F5** or choose **Run > Start Debugging** from the menu to begin debugging. This action automatically launches a local server in a Node window to host your add-in and then automatically opens the host application, such as Excel or Word. This may take several seconds.

1. In the host application, your add-in is now ready to use. Select **Show Taskpane** or run any other add-in command. A dialog box will appear, reading:

   > WebView Stop On Load.
   > To debug the webview, attach VS Code to the webview instance using the Microsoft Debugger for Edge extension, and click OK to continue. To prevent this dialog from appearing in the future, click Cancel.

   Select **OK**.

   > [!NOTE]
   > If you select **Cancel**, the dialog won't be shown again while this instance of the add-in is running. However, if you restart your add-in, you'll see the dialog again.

1. You're now able to set breakpoints in your project's code and debug.

   > [!NOTE]
   > Breakpoints in calls of `Office.initialize` or `Office.onReady` are ignored. For details about these methods, see [Initialize your Office Add-in](../develop/initialize-add-in.md).

> [!IMPORTANT]
> The best way to stop a debugging session is to select **Shift+F5** or choose **Run > Stop Debugging** from the menu. This action should close the Node server window and attempt to close the host application, but there will be a prompt on the host application asking you whether to save the document or not. Make an appropriate choice and let the host application close. Avoid manually closing the Node window or host application. Doing so can cause bugs especially when you are stopping and starting debugging sessions repeatedly.
>
> If debugging stops working; for example, if breakpoints are being ignored; stop debugging. Then, if necessary, close all host application windows and the Node window. Finally, close Visual Studio Code and reopen it.

## See also

- [Test and debug Office Add-ins](test-debug-office-add-ins.md)
- [Microsoft Office Add-in Debugger Extension for Visual Studio Code](debug-with-vs-extension.md)
- [Attach a debugger from the task pane](attach-debugger-from-task-pane.md)
