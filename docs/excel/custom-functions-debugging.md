---
ms.date: 07/10/2020
description: Learn how to debug your Excel custom functions that don't use a task pane.
title: UI-less custom functions debugging
localization_priority: Normal
---
# UI-less custom functions debugging

Debugging for custom functions that don't use a task pane or other user interface elements (UI-less custom functions) can be accomplished by multiple means, depending on what platform you're using.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

On Windows:
- [Excel Desktop and Visual Studio Code (VS Code) debugger](#use-the-vs-code-debugger-for-excel-desktop)
- [Excel on the web and VS Code debugger](#use-the-vs-code-debugger-for-excel-in-microsoft-edge)
- [Excel on the web and browser tools](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [Command line](#use-the-command-line-tools-to-debug)

On Mac:
- [Excel on the web and browser tools](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [Command line](#use-the-command-line-tools-to-debug)

> [!NOTE]
> For simplicity, this article shows debugging in the context of using Visual Studio Code to edit, run tasks, and in some cases use the debug view. If you are using a different editor or command line tool, see the [command line instructions](#commands-for-building-and-running-your-add-in) at the end of this article.

## Requirements

Before starting to debug, you should use the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) to create a custom functions project. For guidance about how to create a custom functions project, see the [custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md).

## Use the VS Code debugger for Excel Desktop

You can use VS Code to debug UI-less custom functions in Office Excel on the desktop.

> [!NOTE]
> Desktop debugging for the Mac is not available but can be achieved [using the browser tools and command line to debug Excel on the web](#use-the-command-line-tools-to-debug)).

### Run your add-in from VS Code

1. Open your custom functions root project folder in [VS Code](https://code.visualstudio.com/).
2. Choose **Terminal > Run Task** and type or select **Watch**. This will monitor and rebuild for any file changes.
3. Choose **Terminal > Run Task** and type or select **Dev Server**.

### Start the VS Code debugger

4. Choose **View > Debug** or enter **Ctrl+Shift+D** to switch to debug view.
5. From the Debug options, choose **Excel Desktop**.
6. Select **F5** (or choose **Debug -> Start Debugging** from the menu) to begin debugging. A new Excel workbook will open with your add-in already sideloaded and ready to use.

### Start debugging

1. In VS Code, open your source code script file (**functions.js** or **functions.ts**).
2. [Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.
3. In the Excel workbook, enter a formula that uses your custom function.

At this point execution will stop on the line of code where you set the breakpoint. Now you can step through your code, set watches, and use any VS Code debugging features you need.

## Use the VS Code debugger for Excel in Microsoft Edge

You can use VS Code to debug UI-less custom functions in Excel on the Microsoft Edge browser. To use VS Code with Microsoft Edge, you must install the [Debugger for Microsoft Edge](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge) extension.

### Run your add-in from VS Code

1. Open your custom functions root project folder in [VS Code](https://code.visualstudio.com/).
2. Choose **Terminal > Run Task** and type or select **Watch**. This will monitor and rebuild for any file changes.
3. Choose **Terminal > Run Task** and type or select **Dev Server**.

### Start the VS Code debugger

4. Choose **View > Debug** or enter **Ctrl+Shift+D** to switch to debug view.
5. From the Debug options, choose **Office Online (Microsoft Edge)**.
6. Open Excel in the Microsoft Edge browser and create a new workbook.
7. Choose **Share** in the ribbon and copy the link for the URL for this new workbook.
8. Select **F5** (or choose **Debug > Start Debugging** from the menu) to begin debugging. A prompt will appear, which asks for the URL of your document.
9. Paste in the URL for your workbook and press Enter.

### Sideload your add-in

1. Select the **Insert** tab on the ribbon and in the **Add-ins** section, choose **Office Add-ins**.
2. On the **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then **Upload My Add-in**.
    
    ![The Office Add-ins dialog with a drop-down in the upper right reading "Manage my add-ins" and a drop-down below it with the option "Upload My Add-in"](../images/office-add-ins-my-account.png)

3. **Browse** to the add-in manifest file and then select **Upload**.
    
    ![The upload add-in dialog with buttons for browse, upload, and cancel.](../images/upload-add-in.png)


### Set breakpoints
1. In VS Code, open your source code script file (**functions.js** or **functions.ts**).
2. [Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.
3. In the Excel workbook, enter a formula that uses your custom function.

## Use the browser developer tools to debug custom functions in Excel on the web

You can use the browser developer tools to debug UI-less custom functions in Excel on the web. The following steps work for both Windows and macOS.

### Run your add-in from Visual Studio Code

1. Open your custom functions root project folder in [Visual Studio Code (VS Code)](https://code.visualstudio.com/).
2. Choose **Terminal > Run Task** and type or select **Watch**. This will monitor and rebuild for any file changes.
3. Choose **Terminal > Run Task** and type or select **Dev Server**.

### Sideload your add-in

1. Open [Office on the web](https://office.live.com/).
2. Open a new Excel workbook.
3. Open the **Insert** tab on the ribbon and, in the **Add-ins** section, choose **Office Add-ins**.
4. On the **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then **Upload My Add-in**.
    
    ![The Office Add-ins dialog with a drop-down in the upper right reading "Manage my add-ins" and a drop-down below it with the option "Upload My Add-in"](../images/office-add-ins-my-account.png)

5. **Browse** to the add-in manifest file, and then select **Upload**.
    
    ![The upload add-in dialog with buttons for browse, upload, and cancel.](../images/upload-add-in.png)

> [!NOTE]
> Once you've sideloaded to the document, it will remain sideloaded each time you open the document.

### Start debugging

1. Open developer tools in the browser. For Chrome and most browsers F12 will open the developer tools.
2. In developer tools, open your source code script file using **Cmd+P** or **Ctrl+P** (**functions.js** or **functions.ts**).
3. [Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code. 

If you need to change the code you can make edits in VS Code and save the changes. Refresh the browser to see the changes loaded.

## Use the command line tools to debug

If you are not using VS Code, you can use the command line (such as bash, or PowerShell) to run your add-in. You'll need to use the browser developer tools to debug your code in Excel on the web. You cannot debug the desktop version of Excel using the command line.

1. From the command line run `npm run watch` to watch for and rebuild when code changes occur.
2. Open a second command line window (the first one will be blocked while running the watch.)

3. If you want to start your add-in in the desktop version of Excel, run the following command
    
    `npm run start:desktop`
    
    Or if you prefer to start your add-in in Excel on the web run the following command
    
    `npm run start:web`
    
    For Excel on the web you also need to sideload your add-in. Follow the steps in [Sideload your add-in](#sideload-your-add-in) to sideload your add-in. Then continue to the next section to start debugging.
    
4. Open developer tools in the browser. For Chrome and most browsers F12 will open the developer tools.
5. In developer tools, open your source code script file (**functions.js** or **functions.ts**). Your custom functions code may be located near the end of the file.
6. In the custom function source code, apply a breakpoint by selecting a line of code.

If you need to change the code you can make edits in Visual Studio and save the changes. Refresh the browser to see the changes loaded.

### Commands for building and running your add-in

There are several build tasks available:
- `npm run watch`: builds for development and automatically rebuilds when a source file is saved
- `npm run build-dev`: builds for development once
- `npm run build`: builds for production
- `npm run dev-server`: runs the web server used for development

You can use the following tasks to start debugging on desktop or online.
- `npm run start:desktop`: Starts Excel on desktop and sideloads your add-in.
- `npm run start:web`: Starts Excel on the web and sideloads your add-in.
- `npm run stop`: Stops Excel and debugging.

## Next steps
Learn about [authentication practices for UI-less custom functions](custom-functions-authentication.md).

## See also

* [Custom functions troubleshooting](custom-functions-troubleshooting.md)
* [Error handling for custom functions in Excel](custom-functions-errors.md)
* [Create custom functions in Excel](custom-functions-overview.md)
