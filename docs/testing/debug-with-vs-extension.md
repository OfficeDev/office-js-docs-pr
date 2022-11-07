---
title: Debug add-ins on Windows using Visual Studio Code and Microsoft Edge legacy WebView (EdgeHTML)
description: Learn how to debug Office Add-ins that use Microsoft Edge Legacy WebView (EdgeHTML) by using the Office Add-in Debugger Extension in VS Code.
ms.date: 02/15/2022
ms.localizationpriority: medium
---

# Microsoft Office Add-in Debugger Extension for Visual Studio Code

Office Add-ins running on Windows can use the Office Add-in Debugger Extension in Visual Studio Code to debug against Microsoft Edge Legacy with the original WebView (EdgeHTML) runtime. 

> [!IMPORTANT]
> This article only applies when Office runs add-ins in the original WebView (EdgeHTML) runtime, as explained in [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md). For instructions about debugging in Visual Studio code against Microsoft Edge WebView2 (Chromium-based), see [Microsoft Office Add-in Debugger Extension for Visual Studio Code](debug-desktop-using-edge-chromium.md).

> [!TIP]
> If you cannot, or don't wish to, debug using tools built into Visual Studio Code; or you are encountering a problem that only occurs when the add-in is run outside Visual Studio Code, you can debug Edge Legacy (EdgeHTML) runtime by using the Edge Legacy developer tools as described in [Debug add-ins using developer tools in Microsoft Edge Legacy](debug-add-ins-using-devtools-edge-legacy.md).

This debugging mode is dynamic, allowing you to set breakpoints while code is running. You can see changes in your code immediately while the debugger is attached, all without losing your debugging session. Your code changes also persist, so you can see the results of multiple changes to your code. The following image shows this extension in action.

![Office Add-in Debugger Extension debugging a section of Excel add-ins.](../images/vs-debugger-extension-for-office-addins.jpg)

## Prerequisites

- [Visual Studio Code](https://code.visualstudio.com/)
- [Node.js (version 10+)](https://nodejs.org/)
- Windows 10, 11
- [Microsoft Edge](https://www.microsoft.com/edge) A combination of platform and Office application that supports Microsoft Edge Legacy with with the original webview (EdgeHTML) as explained in [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).

## Install and use the debugger

These instructions assume you have experience using the command line, understand basic JavaScript, and have created an Office Add-in project before using the [Yeoman generator for Office Add-ins](../develop/yeoman-generator-overview.md). If you haven't done this before, consider visiting one of our tutorials, like this [Excel Office Add-in tutorial](../tutorials/excel-tutorial.md).

1. The first step depends on the project and how it was created.

   - If you want to create a project to experiment with debugging in Visual Studio Code, use the [Yeoman generator for Office Add-ins](../develop/yeoman-generator-overview.md). Use any one of our quick start guides, such as the [Outlook add-in quick start](../quickstarts/outlook-quickstart.md), in order to do this.
   - If you want to debug an existing project that was created with Yo Office, skip to the next step.
   - If you want to debug an existing project that was not created with Yo Office, carry out the procedure in the [Appendix](#appendix) and then return to the next step of this procedure.


1. Open VS Code and open your project in it.

1. Within VS Code, select **Ctrl+Shift+X** to open the Extensions bar. Search for the "Microsoft Office Add-in Debugger" extension and install it.

1. Choose  **View > Run** or enter **Ctrl+Shift+D** to switch to debug view.

1. From the **RUN AND DEBUG** options, choose the Edge Legacy option for your host application, such as **Outlook Desktop (Edge Legacy)**. Select **F5** or choose **Run > Start Debugging** from the menu to begin debugging. This action automatically launches a local server in a Node window to host your add-in and then automatically opens the host application, such as Excel or Word. This may take several seconds.

1. In the host application, your add-in is now ready to use. Select **Show Taskpane** or run any other add-in command. A dialog box will appear similar to the following:

   > WebView Stop On Load.
   > To debug the WebView, attach VS Code to the WebView instance using the Microsoft Debugger for Edge extension, and click **OK** to continue. To prevent this dialog from appearing in the future, click **Cancel**.

   Select **OK**.

   > [!NOTE]
   > If you select **Cancel**, the dialog won't be shown again while this instance of the add-in is running. However, if you restart your add-in, you'll see the dialog again.

1. Set a breakpoint in your project's task pane file. To set breakpoints in Visual Studio Code, hover next to a line of code and select the red circle that appears.

    ![Red circle appears on a line of code in Visual Studio Code.](../images/set-breakpoint.jpg)

1. Run functionality in your add-in that calls the lines with breakpoints. You'll see that breakpoints have been hit and you can inspect local variables.

   > [!NOTE]
   > Breakpoints in calls of `Office.initialize` or `Office.onReady` are ignored. For details about these methods, see [Initialize your Office Add-in](../develop/initialize-add-in.md).

> [!IMPORTANT]
> The best way to stop a debugging session is to select **Shift+F5** or choose **Run** > **Stop Debugging** from the menu. This action should close the Node server window and attempt to close the host application, but there will be a prompt on the host application asking you whether to save the document or not. Make an appropriate choice and let the host application close. Avoid manually closing the Node window or host application. Doing so can cause bugs especially when you are stopping and starting debugging sessions repeatedly.
>
> If debugging stops working; for example, if breakpoints are being ignored; stop debugging. Then, if necessary, close all host application windows and the Node window. Finally, close Visual Studio Code and reopen it.

### Appendix

If your project was not created with Yo Office, you need to create a debug configuration for Visual Studio Code.

1. Create a file named `launch.json` in the `\.vscode` folder of the project if there isn't one there already.
1. Ensure that the file has a `configurations` array. The following is a simple example of a `launch.json`.

    ```json
    {
      // Other properties may be here.

      "configurations": [

        // Configuration objects may be here.

      ]

      // Other properties may be here.
    }
    ```

1. Add the following object to the `configurations` array.

    ```json
    {
      "name": "HOST Desktop (Edge Legacy)",
      "type": "office-addin",
      "request": "attach",
      "url": "https://localhost:3000/taskpane.html?_host_Info=HOST$Win32$16.01$en-US$$$$0",
      "port": 9222,
      "timeout": 600000,
      "webRoot": "${workspaceRoot}",
      "preLaunchTask": "Debug: HOST Desktop",
      "postDebugTask": "Stop Debug"
    }
    ```

1. Replace the placeholder `HOST` in all three places with the name of the Office application that the add-in runs in; for example, `Outlook` or `Word`.
1. Save and close the file.

## See also

- [Test and debug Office Add-ins](test-debug-office-add-ins.md)
- [Debug add-ins on Windows using Visual Studio Code and Microsoft Edge WebView2 (Chromium-based)](debug-desktop-using-edge-chromium.md).
- [Debug add-ins using developer tools for Internet Explorer](debug-add-ins-using-f12-tools-ie.md)
- [Debug add-ins using developer tools for Edge Legacy](debug-add-ins-using-devtools-edge-legacy.md)
- [Debug add-ins using developer tools in Microsoft Edge (Chromium-based)](debug-add-ins-using-devtools-edge-chromium.md)
- [Attach a debugger from the task pane](attach-debugger-from-task-pane.md)
- [Runtimes in Office Add-ins](runtimes.md)