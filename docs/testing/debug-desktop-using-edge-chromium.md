---
title: Debug Office Add-ins on Windows using Visual Studio Code and Microsoft Edge
description: Learn how to debug Office Add-ins that use Microsoft Edge in VS Code.
ms.date: 11/06/2025
ms.localizationpriority: high
---

# Debug Office Add-ins on Windows using Visual Studio Code and Microsoft Edge WebView2

Office Add-ins running on Windows can debug against the Edge runtime directly in Visual Studio Code.

> [!TIP]
> If you can't, or don't wish to, debug using tools built into Visual Studio Code; or you're encountering a problem that only occurs when the add-in is run outside Visual Studio Code, you can debug the Microsoft Edge WebView2 runtime by using the Edge developer tools as described in [Debug add-ins using developer tools for Microsoft Edge WebView2](debug-add-ins-using-devtools-edge-chromium.md).

This debugging mode is dynamic, allowing you to set breakpoints while code is running. See changes in your code immediately while the debugger is attached, all without losing your debugging session. Your code changes also persist, so you see the results of multiple changes to your code. The following image shows this extension in action.

:::image type="content" source="../images/vs-debugger-extension-for-office-addins.jpg" alt-text="Office Add-in Debugger Extension debugging a section of Excel add-ins.":::

## Prerequisites

- [Visual Studio Code](https://code.visualstudio.com/)
- [Node.js (version 10+)](https://nodejs.org/)
- Windows 10, 11
- A combination of platform and Office application that supports Microsoft Edge with WebView2 as explained in [Browsers and webview controls used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md). If your version of Office from a Microsoft 365 subscription is earlier than Version 2101, you'll need to install WebView2. For instructions to install WebView2, see [Microsoft Edge WebView2 / Embed web content ... with Microsoft Edge WebView2](https://developer.microsoft.com/microsoft-edge/webview2/).

## Debug a project created with Yo Office

These instructions assume you have experience using the command line, understand basic JavaScript, and have created an Office Add-in project before using the [Yeoman generator for Office Add-ins](../develop/yeoman-generator-overview.md). If you haven't done this before, consider visiting one of our tutorials, such as the [Excel Office Add-in tutorial](../tutorials/excel-tutorial.md).

1. The first step depends on the project and how it was created.

   - If you want to create a project to experiment with debugging in Visual Studio Code, use the [Yeoman generator for Office Add-ins](../develop/yeoman-generator-overview.md). Follow any of the Yo Office quick start guides, such as the [Outlook add-in quick start](../quickstarts/outlook-quickstart-yo.md).
   - If you want to debug an existing project that was created with Yo Office, skip to the next step.

1. Open VS Code and open your project in it.

1. Choose **View** > **Run** or enter <kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>D</kbd> to switch to debug view.

1. From the **RUN AND DEBUG** options, choose the Microsoft Edge option for your host application, such as **Outlook Desktop (Edge Chromium)**. Select <kbd>F5</kbd> or choose **Run** > **Start Debugging** from the menu to begin debugging. This action automatically launches a local server in a Node window to host your add-in and then automatically opens the host application, such as Excel or Word. This may take several seconds.

   > [!TIP]
   > If you aren't using a project created with Yo Office, you may be prompted to adjust a registry key. While in the root folder of your project, run the following in the command line.
   >
   > ```command line
   > npx office-addin-debugging start <your manifest path>
   > ```

   > [!IMPORTANT]
   > If your project was created with older versions of Yo Office, you may see the following error dialog box about 10 - 30 seconds after you start debugging (at which point you may have already gone on to another step in this procedure) and it may be hidden behind the dialog box described in the next step.
   >
   > :::image type="content" source="../images/configured-debug-type-error.jpg" alt-text="Error that says Configured debug type edge is not supported.":::
   >
   > Complete the tasks in the [Appendix](#appendix) and then restart this procedure.

1. In the host application, your add-in is now ready to use. Select **Show Taskpane** or run any other add-in command. A dialog box will appear with text similar to the following:

   > WebView Stop On Load.
   > To debug the webview, attach VS Code to the webview instance using the Microsoft Debugger for Edge extension, and click OK to continue. To prevent this dialog from appearing in the future, click Cancel.

   Select **OK**.

   [!INCLUDE [Cancelling the WebView Stop On Load dialog box](../includes/webview-stop-on-load-cancel-dialog.md)]

1. You're now able to set breakpoints in your project's code and debug. To set breakpoints in Visual Studio Code, hover next to a line of code and select the red circle that appears.

   :::image type="content" source="../images/set-breakpoint.jpg" alt-text="Red circle appears on a line of code in Visual Studio Code.":::

1. Run functionality in your add-in that calls the lines with breakpoints. You'll see that breakpoints have been hit and you can inspect local variables.

   > [!NOTE]
   > Breakpoints in calls of `Office.initialize` or `Office.onReady` are ignored. For details about these functions, see [Initialize your Office Add-in](../develop/initialize-add-in.md).

> [!IMPORTANT]
> The best way to stop a debugging session is to select <kbd>Shift</kbd>+<kbd>F5</kbd> or choose **Run** > **Stop Debugging** from the menu. This action should close the Node server window and attempt to close the host application, but there'll be a prompt on the host application asking you whether to save the document or not. Make an appropriate choice and let the host application close. Avoid manually closing the Node window or host application. Doing so can cause bugs especially when you are stopping and starting debugging sessions repeatedly.
>
> If debugging stops working---for example, if breakpoints are being ignored---stop debugging. Then, if necessary, close all host application windows and the Node window. Finally, close Visual Studio Code and reopen it.

## Debug a project not created with Yo Office

If your project wasn't created with Yo Office, you need to create a debug configuration for Visual Studio Code.

### Configure package.json file

1. Ensure you have a `package.json` file. If you don't already have a package.json file, run `npm init` in the root folder of your project and answer the prompts.
1. Run `npm install office-addin-debugging`. This package sideloads your add-in for debugging.
1. Open the `package.json` file. In the `scripts` section, add the following script.

   ```json
   "start:desktop": "office-addin-debugging start $MANIFEST_FILE$ desktop",
   "dev-server": "$SERVER_START$"
   ```

1. Replace `$MANIFEST_FILE$` with the correct file name and folder location of your manifest.
1. Replace `$SERVER_START$` with the command to start your web server. Later in these steps, the `office-addin-debugging` package will specifically look for the `dev-server` script to launch your web server.
1. Save and close the `package.json` file.

### Configure launch.json file

1. Create a file named `launch.json` in the `\.vscode` folder of the project if there isn't one there already.
1. Copy the following JSON into the file.

   ```json
   {
     // Other properties may be here.
     "configurations": [
       {
         "name": "$HOST$ Desktop (Edge Chromium)",
         "type": "msedge",
         "request": "attach",
         "useWebView": true,
         "port": 9229,
         "timeout": 600000,
         "webRoot": "${workspaceRoot}",
         "preLaunchTask": "Debug: Excel Desktop"
       }
     ]
     // Other properties may be here.
   }
   ```

   > [!NOTE]
   > If you already have a `launch.json` file, just add the single configuration to the `configurations` section.

1. Replace the placeholder `$HOST$` with the name of the Office application that the add-in runs in. For example, `Outlook` or `Word`.
1. Save and close the file.

### Configure tasks.json

1. Create a file named `tasks.json` in the `\.vscode` folder of the project.
1. Copy the following JSON into the file. It contains a task that starts debugging for your add-in.

   ```json
   {
     "version": "2.0.0",
     "tasks": [
       {
         "label": "Debug: $HOST$ Desktop",
         "type": "shell",
         "command": "npm",
         "args": ["run", "start:desktop", "--", "--app", "$HOST$"],
         "presentation": {
           "clear": true,
           "panel": "dedicated"
         },
         "problemMatcher": []
       }
     ]
   }
   ```

   > [!NOTE]
   > If you already have a `tasks.json` file, just add the single task to the `tasks` section.

1. Replace both instances of the placeholder `$HOST$` with the name of the Office application that the add-in runs in. For example, `Outlook` or `Word`.

You can now debug your project using the VS Code debugger (F5).

### Appendix

1. In the error dialog box, select the **Cancel** button.
1. If debugging doesn't stop automatically, select <kbd>Shift</kbd>+<kbd>F5</kbd> or choose **Run** > **Stop Debugging** from the menu.
1. Close the Node window where the local server is running, if it doesn't close automatically.
1. Close the Office application if it doesn't close automatically.
1. Open the `\.vscode\launch.json` file in the project.
1. In the `configurations` array, there are several configuration objects. Find the one whose name has the pattern `$HOST$ Desktop (Edge Chromium)`, where $HOST$ is an Office application that your add-in runs in; for example, `Outlook Desktop (Edge Chromium)` or `Word Desktop (Edge Chromium)`.
1. Change the value of the `"type"` property from `"edge"` to `"pwa-msedge"`.
1. Change the value of the `"useWebView"` property from the string `"advanced"` to the boolean `true` (note there are no quotation marks around the `true`).
1. Save the file.
1. Close VS Code.

## See also

- [Test and debug Office Add-ins](test-debug-office-add-ins.md)
- [Debug add-ins using developer tools in Microsoft Edge](debug-add-ins-using-devtools-edge-chromium.md)
- [Attach a debugger from the task pane](attach-debugger-from-task-pane.md)
- [Runtimes in Office Add-ins](runtimes.md)
