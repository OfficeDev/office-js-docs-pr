---
title: Custom functions debugging in a non-shared runtime
description: Learn how to debug your Excel custom functions that don't use a shared runtime.
ms.date: 07/11/2022
ms.localizationpriority: medium
---

# Custom functions debugging

This article discusses debugging only for custom functions that **don't use a [shared runtime](../testing/runtimes.md#shared-runtime)**. To debug custom functions add-ins that use a shared runtime, see [Configure your Office Add-in to use a shared runtime: Debug](../develop/configure-your-add-in-to-use-a-shared-runtime.md#debug).

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

> [!TIP]
> This debugging process doesn't work with projects that are created with the **Office Add-in project containing the manifest only** option in the Yeoman generator. The scripts that are referred to later in this article aren't installed with that option. To debug an add-in that is created with this option, see the instructions in one of the following articles, as appropriate.
>
> - [Debug add-ins using developer tools in Microsoft Edge (Chromium-based)](../testing/debug-add-ins-using-devtools-edge-chromium.md)
> - [Debug add-ins using developer tools in Internet Explorer](../testing/debug-add-ins-using-f12-tools-ie.md)
> - [Debug Office Add-ins on a Mac](../testing/debug-office-add-ins-on-ipad-and-mac.md)

The process of debugging a custom function for add-ins that don't use a shared runtime varies depending on the target platform (Windows, Mac, or web), whether you are using Visual Studio Code or a different IDE, and the operating system of your development computer. Use the links in the following table to visit sections of this article that are relevant to your debugging scenario. In this table, "CF-NSR" refers to custom functions in a non-shared runtime.

| **Target platform** | **Visual Studio Code** | **Other IDE** |
|--------------|-------------|-------------|
| Excel on Windows | [Use the VS Code debugger for Excel on Windows](#use-the-vs-code-debugger-for-excel-on-windows) | Debugging CF-NSR outside VS Code isn't supported. Debug against Excel on the web. |
| Excel on the web | Windows development computer: [Use the VS Code debugger for Excel in Microsoft Edge](#use-the-vs-code-debugger-for-excel-in-microsoft-edge)</br>Mac or Windows development computer: [Use VS Code and the browser development tools](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web) | [Use the command line tools](#use-the-command-line-tools-to-debug)|
| Excel on Mac |  VS Code debugging of CF-NSR isn't supported. Debug against Excel on the web. | [Use the command line tools](#use-the-command-line-tools-to-debug)|

> [!NOTE]
> This article primarily shows debugging in the context of using Visual Studio Code to edit, run tasks, and use the debug view. If you're using a different editor or command line tool, see [Commands for building and running your add-in](#commands-for-building-and-running-your-add-in) at the end of this article.

## Use the VS Code debugger for Excel on Windows

You can use VS Code to debug custom functions that don't use a shared runtime in Office Excel on the desktop.

> [!IMPORTANT]
> There is a known issue with the following debugging steps. The steps work for a project installed with the **Excel Custom Functions Add-in project** option in the Yeoman generator with **TypeScript** selected as the script type, but the steps do not work for a project installed with **JavaScript** selected as the script type. For additional information, see [OfficeDev/office-js-docs-pr issue #3355](https://github.com/OfficeDev/office-js-docs-pr/issues/3355).

### Run your add-in from VS Code

1. Open your custom functions root project folder in [VS Code](https://code.visualstudio.com/).
1. Choose **Terminal > Run Task** and type or select **Watch**. This will monitor and rebuild for any file changes.
1. Choose **Terminal > Run Task** and type or select **Dev Server**.

### Start the VS Code debugger

1. Choose **View > Run** or enter **Ctrl+Shift+D** to switch to debug view.
1. From the **Run and Debug** drop-down menu, choose **Excel Desktop (Custom Functions)**.

    :::image type="content" source="../images/custom-functions-run-and-debug-menu.jpg" alt-text="A screenshot showing Excel Desktop (Custom Functions) in the Run and Debug drop-down menu.":::

1. Select **F5** (or select **Run -> Start Debugging** from the menu) to begin debugging. A new Excel workbook will open with your add-in already sideloaded and ready to use.

### Start debugging

1. In VS Code, open your source code script file (**functions.js** or **functions.ts**).
2. [Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.
3. In the Excel workbook, enter a formula that uses your custom function.

At this point, execution will stop on the line of code where you set the breakpoint. Now you can step through your code, set watches, and use any VS Code debugging features you need.

## Use the VS Code debugger for Excel in Microsoft Edge

You can use VS Code to debug custom functions that don't use a shared runtime in Excel on the Microsoft Edge browser. To use VS Code with Microsoft Edge, you must install the [Microsoft Edge DevTools extension for Visual Studio Code](/microsoft-edge/visual-studio-code/microsoft-edge-devtools-extension).

### Run your add-in from VS Code

1. Open your custom functions root project folder in [VS Code](https://code.visualstudio.com/).
1. Choose **Terminal > Run Task** and type or select **Watch**. This will monitor and rebuild for any file changes.
1. Choose **Terminal > Run Task** and type or select **Dev Server**.

### Start the VS Code debugger

1. Choose **View > Run** or enter **Ctrl+Shift+D** to switch to debug view.
1. From the Debug options, choose **Office Online (Edge Chromium)**.
1. Open Excel in the Microsoft Edge browser and create a new workbook.
1. Choose **Share** in the ribbon and copy the link for the URL for this new workbook.
1. Select **F5** (or select **Run > Start Debugging** from the menu) to begin debugging. A prompt will appear, which asks for the URL of your document.
1. Paste in the URL for your workbook and press Enter.

### Sideload your add-in

1. Select the **Insert** tab on the ribbon and in the **Add-ins** section, choose **Office Add-ins**.
2. On the **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then **Upload My Add-in**.
  
    ![The Office Add-ins dialog with a drop-down in the upper right reading "Manage my add-ins" and a drop-down below it with the option "Upload My Add-in".](../images/office-add-ins-my-account.png)

3. **Browse** to the add-in manifest file and then select **Upload**.
  
    ![The upload add-in dialog with buttons for browse, upload, and cancel.](../images/upload-add-in.png)

### Set breakpoints

1. In VS Code, open your source code script file (**functions.js** or **functions.ts**).
2. [Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.
3. In the Excel workbook, enter a formula that uses your custom function.

## Use the browser developer tools to debug custom functions in Excel on the web

You can use the browser developer tools to debug custom functions that don't use a shared runtime in Excel on the web. The following steps work for both Windows and macOS.

### Run your add-in from Visual Studio Code

1. Open your custom functions root project folder in [Visual Studio Code (VS Code)](https://code.visualstudio.com/).
2. Choose **Terminal > Run Task** and type or select **Watch**. This will monitor and rebuild for any file changes.
3. Choose **Terminal > Run Task** and type or select **Dev Server**.

### Sideload your add-in

1. Open [Office on the web](https://office.live.com/).
2. Open a new Excel workbook.
3. Open the **Insert** tab on the ribbon and, in the **Add-ins** section, choose **Office Add-ins**.
4. On the **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then **Upload My Add-in**.
  
    ![The Office Add-ins dialog with a drop-down in the upper right reading "Manage my add-ins" and a drop-down below it with the option "Upload My Add-in".](../images/office-add-ins-my-account.png)

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

If you aren't using VS Code, you can use the command line (such as bash, or PowerShell) to run your add-in. You'll need to use the browser developer tools to debug your code in Excel on the web. You cannot debug the desktop version of Excel using the command line.

1. From the command line run `npm run watch` to watch for and rebuild when code changes occur.
2. Open a second command line window (the first one will be blocked while running the watch.)

3. If you want to start your add-in in the desktop version of Excel, run the following command.
  
    `npm run start:desktop`
  
    Or if you prefer to start your add-in in Excel on the web run the following command.
  
    `npm run start:web -- --document {url}` (where `{url}` is the URL of an Excel file on OneDrive or SharePoint)
  
    If your add-in doesn't sideload in the document, follow the steps in [Sideload your add-in](#sideload-your-add-in) to sideload your add-in. Then continue to the next section to start debugging.
  
4. Open developer tools in the browser. For Chrome and most browsers F12 will open the developer tools.
5. In developer tools, open your source code script file (**functions.js** or **functions.ts**). Your custom functions code may be located near the end of the file.
6. In the custom function source code, apply a breakpoint by selecting a line of code.

If you need to change the code you can make edits in Visual Studio and save the changes. Refresh the browser to see the changes loaded.

### Commands for building and running your add-in

There are several build tasks available.

- `npm run watch`: builds for development and automatically rebuilds when a source file is saved
- `npm run build-dev`: builds for development once
- `npm run build`: builds for production
- `npm run dev-server`: runs the web server used for development

You can use the following tasks to start debugging on desktop or online.

- `npm run start:desktop`: Starts Excel on desktop and sideloads your add-in.
- `npm run start:web -- --document {url}` (where `{url}` is the URL of an Excel file on OneDrive or SharePoint): Starts Excel on the web and sideloads your add-in.
- `npm run stop`: Stops Excel and debugging.

## Next steps

Learn about [Authentication for custom functions without a shared runtime](custom-functions-authentication.md).

## See also

- [Custom functions troubleshooting](custom-functions-troubleshooting.md)
- [Error handling for custom functions in Excel](custom-functions-errors.md)
- [Create custom functions in Excel](custom-functions-overview.md)
- [JavaScript-only runtime](../testing/runtimes.md#javascript-only-runtime)
