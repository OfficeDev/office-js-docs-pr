---
ms.date: 05/03/2019
description: Debug your custom functions in Excel.
title: Custom functions debugging (preview)
localization_priority: Normal
---
# Custom functions debugging (preview)

Debugging for custom functions can be accomplished by multiple means, depending on what platform you're using.

On Windows:
- [Excel Desktop and Visual Studio Code (VS Code) debugger](#use-the-vs-code-debugger-for-excel-desktop)
- [Excel Online and VS Code debugger](#use-the-vs-code-debugger-for-excel-online-in-microsoft-edge)
- [Excel Online and browser tools](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-online)
- [Command line](#use-the-command-line-tools-to-debug)

On Mac:
- [Excel Online and browser tools](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-online)
- [Command line](#use-the-command-line-tools-to-debug)

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

> [!NOTE]
> For simplicity, this article shows debugging in the context of using Visual Studio Code to edit, run tasks, and in some cases use the debug view. If you are using a different editor or command line tool, see the [command line instructions](#use-the-command-line-tools-to-debug) at the end of this article.

## Requirements

Before starting to debug, you should create a custom functions add-in project. You can do this using the [custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md). For instructions on trusting certificates, see [Adding self-signed certificates as trusted root certificates](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).

## Use the VS Code debugger for Excel Desktop

You can use VS Code to debug custom functions in Office Excel on the desktop.

> [!NOTE]
> Desktop debugging for the Mac is not available but can be achieved [using the browser tools to debug Excel Online](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-online).

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

## Use the VS Code debugger for Excel Online in Microsoft Edge

You can use VS Code to debug custom functions in Excel Online in the Microsoft Edge browser. To use VS Code with Microsoft Edge, you must install the [Debugger for Microsoft Edge](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge) extension.

### Run your add-in from VS Code

1. Open your custom functions root project folder in [VS Code](https://code.visualstudio.com/).
2. Choose **Terminal > Run Task** and type or select **Watch**. This will monitor and rebuild for any file changes.
3. Choose **Terminal > Run Task** and type or select **Dev Server**. 

### Start the VS Code debugger

4. Choose **View > Debug** or enter **Ctrl+Shift+D** to switch to debug view.
5. From the Debug options, choose **Office Online (Edge)**.
6. Open Excel Online using the Microsoft Edge browser, open Excel Online, and create a new workbook.
7. Choose **Share** in the ribbon and copy the link for the URL for this new workbook.
8. Select **F5** (or choose **Debug > Start Debugging** from the menu) to begin debugging. A prompt will appear, which asks for the URL of your document.
9. Paste in the URL for your workbook and press Enter.

### Sideload your add-in   

1. Select the  **Insert** tab on the ribbon and in the **Add-ins** section, choose **Office Add-ins**.
2. On the  **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then  **Upload My Add-in**.
    
    ![The Office Add-ins dialog with a drop-down in the upper right reading "Manage my add-ins" and a drop-down below it with the option "Upload My Add-in"](../images/office-add-ins-my-account.png)

3.  **Browse** to the add-in manifest file and then select **Upload**.
    
    ![The upload add-in dialog with buttons for browse, upload, and cancel.](../images/upload-add-in.png)


### Set breakpoints
1. In VS Code, open your source code script file (**functions.js** or **functions.ts**).
2. [Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.
3. In the Excel workbook, enter a formula that uses your custom function.

## Use the browser developer tools to debug custom functions in Excel Online

You can use the browser developer tools to debug custom functions in Excel Online. The following steps work for both Windows and macOS.

### Run your add-in from Visual Studio Code

1. Open your custom functions root project folder in [Visual Studio Code (VS Code)](https://code.visualstudio.com/).
2. Choose **Terminal > Run Task** and type or select **Watch**. This will monitor and rebuild for any file changes.
3. Choose **Terminal > Run Task** and type or select **Dev Server**. 

### Sideload your add-in   

1. Open [Microsoft Office Online](https://office.live.com/).
2. Open a new Excel workbook.
3. Open the  **Insert** tab on the ribbon and, in the **Add-ins** section, choose **Office Add-ins**.
4. On the  **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then  **Upload My Add-in**.
    
    ![The Office Add-ins dialog with a drop-down in the upper right reading "Manage my add-ins" and a drop-down below it with the option "Upload My Add-in"](../images/office-add-ins-my-account.png)

5.  **Browse** to the add-in manifest file, and then select **Upload**.
    
    ![The upload add-in dialog with buttons for browse, upload, and cancel.](../images/upload-add-in.png)

> [!NOTE]
> Once you've sideloaded to the document, it will remain sideloaded each time you open the document.

### Start debugging

1. Open developer tools in the browser. For Chrome and most browsers F12 will open the developer tools.
2. In developer tools, open your source code script file using **Cmd+P** or **Ctrl+P** (**functions.js** or **functions.ts**).
3. [Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code. 

If you need to change the code you can make edits in VS Code and save the changes. Refresh the browser to see the changes loaded.

## Use the command line tools to debug

If you are not using VS Code, you can use the command line (such as bash, or PowerShell) to run your add-in. You'll need to use the browser developer tools to debug your code in Excel Online. You cannot debug the desktop version of Excel using the command line.

1. From the command line run `npm run watch` to watch for and rebuild when code changes occur.
2. Open a second command line window (the first one will be blocked while running the watch.)

3. If you want to start your add-in in the desktop version of Excel, run the following command
    
    `npm run start:desktop`
    
    Or if you prefer to start your add-in in Excel Online run the following command
    
    `npm run start:web`
    
    For Excel Online you also need to sideload your add-in. Follow the steps in [Sideload your add-in](#sideload-your-add-in) to sideload your add-in. Then continue to the next section to start debugging.
    
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
- `npm run start:web`: Starts Excel Online and sideloads your add-in.
- `npm run stop`: Stops Excel and debugging.

## Next steps
Learn about [authentication practices in custom functions](custom-functions-authentication.md). Or, review [custom function's unique architecture](custom-functions-architecture.md).

## See also

* [Custom functions troubleshooting](custom-functions-troubleshooting.md)
* [Custom functions best practices](custom-functions-best-practices.md)
* [Custom functions changelog](custom-functions-changelog.md)
* [Make your custom functions compatible with XLL user-defined functions](make-custom-functions-compatible-with-xll-udf.md)
* [Create custom functions in Excel](custom-functions-overview.md)
