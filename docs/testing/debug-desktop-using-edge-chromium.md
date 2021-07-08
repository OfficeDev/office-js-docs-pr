---
title: Debug add-ins on Windows using Microsoft Edge WebView2 (Chromium-based)
description: 'Learn how to debug Office Add-ins that use Microsoft Edge WebView2 (Chromium-based) by using the Debugger for Microsoft Edge extension in VS Code.'
ms.date: 01/29/2021
localization_priority: Priority
---
# Debug add-ins on Windows using Edge Chromium WebView2

Office Add-ins running on Windows can use the Debugger for Microsoft Edge extension in VS Code to debug against the Edge Chromium WebView2 runtime.

## Prerequisites

- [Visual Studio Code](https://code.visualstudio.com/) (must be run as an administrator)
- [Node.js (version 10+)](https://nodejs.org/)
- Windows 10
- [Microsoft Edge Chromium available to Windows Insiders](https://www.microsoftedgeinsider.com/)

## Install and use the debugger

1. Create a project using the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office). You can use any one of our quick start guides, such as the [Outlook add-in quickstart](../quickstarts/outlook-quickstart.md), in order to do this.

    > [!TIP]
    > If you aren't using a Yeoman generator based add-in, you need to adjust a registry key. While in the root folder of your project, run the following in the command line: `office-add-in-debugging start <your manifest path>`.

1. Open your project in VS Code. Within VS Code, select **CTRL + SHIFT + X** to open the Extensions bar. Search for the "Debugger for Microsoft Edge" extension and install it.

1. In the **.vscode** folder of your project, open the **launch.json** file. Add the following code to the configurations section.

      ```JSON
        {
          "name": "Debug Office Add-in (Edge Chromium)",
          "type": "edge",
          "request": "attach",
          "useWebView": "advanced",
          "port": 9229,
          "timeout": 600000,
          "webRoot": "${workspaceRoot}",
        },
      ```

1. Next, choose  **View > Debug** or enter **CTRL + SHIFT + D** to switch to debug view.

1. From the Debug options, choose the Edge Chromium option for your host application, such as **Excel Desktop (Edge Chromium)**. Select **F5** or choose **Debug > Start Debugging** from the menu to begin debugging.

1. In the host application, such as Excel, your add-in is now ready to use. Select **Show Taskpane** or run any other add-in command. A dialog box will appear, reading:

    > WebView Stop On Load.
    > To debug the webview, attach VS Code to the webview instance using the Microsoft Debugger for Edge extension, and click OK to continue. To prevent this dialog from appearing in the future, click Cancel."

    Select **OK**.

    > [!NOTE]
    > If you select **Cancel**, the dialog won't be shown again while this instance of the add-in is running. However, if you restart your add-in, you'll see the dialog again.

1. You're now able to set breakpoints in your project's code and debug.

## See also

- [Test and debug Office Add-ins](test-debug-office-add-ins.md)
- [Microsoft Office Add-in Debugger Extension for Visual Studio Code](debug-with-vs-extension.md)
- [Attach a debugger from the task pane](attach-debugger-from-task-pane.md)