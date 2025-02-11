---
title: Custom functions debugging in a non-shared runtime
description: Learn how to debug your Excel custom functions that don't use a shared runtime.
ms.date: 01/03/2024
ms.topic: troubleshooting
ms.localizationpriority: medium
---

# Custom functions debugging in a non-shared runtime

This article discusses debugging only for custom functions that **don't use a [shared runtime](../testing/runtimes.md#shared-runtime)**. To debug custom functions add-ins that use a shared runtime, see [Overview of debugging Office Add-ins](../testing/debug-add-ins-overview.md).

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

> [!TIP]
> The debugging techniques that are described in this article don't work with projects that are created with the **Office Add-in project containing the manifest only** option in the Yeoman generator. The scripts that are referred to later in this article aren't installed with that option. To debug an add-in that is created with this option, see the instructions in one of the following articles, as appropriate.
>
> - [Debug add-ins using developer tools in Microsoft Edge (Chromium-based)](../testing/debug-add-ins-using-devtools-edge-chromium.md)
> - [Debug add-ins using developer tools in Internet Explorer](../testing/debug-add-ins-using-f12-tools-ie.md)
> - [Debug Office Add-ins on a Mac](../testing/debug-office-add-ins-on-ipad-and-mac.md)

The process of debugging a custom function for add-ins that don't use a shared runtime varies depending on the target platform (Windows, Mac, or web) and on whether you are using Visual Studio Code or a different IDE. Use the links in the following table to visit sections of this article that are relevant to your debugging scenario. In this table, "CF-NSR" refers to custom functions in a non-shared runtime.

| Target platform | Visual Studio Code | Other IDE |
|--------------|-------------|-------------|
| Excel on the web | [Use VS Code and the browser development tools](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web) | [Use the command line tools](#use-the-command-line-tools-to-debug)|
| Excel on Windows | [Use VS Code and the browser development tools](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web) | Debugging CF-NSR that are running in Excel on Windows outside VS Code isn't supported. Debug against Excel on the web. |
| Excel on Mac |  VS Code debugging of CF-NSR that are running in Excel on Mac isn't supported. Debug against Excel on the web. | [Use the command line tools](#use-the-command-line-tools-to-debug)|

## Use the browser developer tools to debug custom functions in Excel on the web

You can use the browser developer tools to debug custom functions that don't use a shared runtime in Excel on the web. The following steps work for both Windows and macOS.

### Run your add-in from Visual Studio Code

1. Open your custom functions root project folder in [Visual Studio Code (VS Code)](https://code.visualstudio.com/).
1. Choose **Terminal** > **Run Task** and type or select **Watch**. This will monitor and rebuild for any file changes.
1. Choose **Terminal** > **Run Task** and type or select **Dev Server**.

### Sideload your add-in

1. Open [Office on the web](https://office.live.com/).
1. Open a new Excel workbook.
1. Select **Home** > **Add-ins**, then select **More Settings**.
1. On the **Office Add-ins** dialog, select **Upload My Add-in**.
1. **Browse** to the add-in manifest file, and then select **Upload**.
  
    ![The upload add-in dialog with buttons for browse, upload, and cancel.](../images/upload-add-in.png)

> [!NOTE]
> Once you've sideloaded to the document, it will remain sideloaded each time you open the document.

### Start debugging

1. Open developer tools in the browser. For Chrome and most browsers F12 will open the developer tools.
1. In developer tools, open your source code script file using <kbd>Cmd</kbd>+<kbd>P</kbd> or <kbd>Ctrl</kbd>+<kbd>P</kbd> (**functions.js** or **functions.ts**).
1. [Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.

If you need to change the code you can make edits in VS Code and save the changes. Refresh the browser to see the changes loaded.

## Use the command line tools to debug

If you aren't using VS Code, you can use the command line (such as bash, or PowerShell) to run your add-in. You'll need to use the browser developer tools to debug your code in Excel on the web. You cannot debug the desktop version of Excel using the command line.

1. From the command line run `npm run watch` to watch for and rebuild when code changes occur.
1. Open a second command line window (the first one will be blocked while running the watch.)

1. If you want to start your add-in in the desktop version of Excel and the "scripts" section of the project's package.json file has a "start:desktop" script, then run `npm run start:desktop`; otherwise, run `npm run start`.
  
    Or if you prefer to start your add-in in Excel on the web run the following command.
  
    `npm run start -- web --document {url}` (where `{url}` is the URL of an Excel file on OneDrive or SharePoint)
  
    [!include[Mac command line note](../includes/mac-command-line.md)]
  
    If your add-in doesn't sideload in the document, follow the steps in [Sideload your add-in](#sideload-your-add-in) to sideload your add-in. Then continue to the next section to start debugging.
  
1. Open developer tools in the browser. For Chrome and most browsers F12 will open the developer tools.
1. In developer tools, open your source code script file (**functions.js** or **functions.ts**). Your custom functions code may be located near the end of the file.
1. In the custom function source code, apply a breakpoint by selecting a line of code.

If you need to change the code, you can make edits in VS Code and save the changes. Refresh the browser to see the changes loaded.

### Commands for building and running your add-in

There are several build tasks available.

- `npm run watch`: builds for development and automatically rebuilds when a source file is saved
- `npm run build-dev`: builds for development once
- `npm run build`: builds for production
- `npm run dev-server`: runs the web server used for development

You can use the following tasks to start debugging on desktop or online.

- `npm run start:desktop`: Starts Excel on desktop and sideloads your add-in. If the "start:desktop" script isn't present in the "scripts" section of the project's package.json file, then run `npm run start` instead.
- `npm run start -- web --document {url}` (where `{url}` is the URL of an Excel file on OneDrive or SharePoint): Starts Excel on the web and sideloads your add-in.

  [!include[Mac command line note](../includes/mac-command-line.md)]

- `npm run stop`: Stops Excel and debugging.

## Next steps

Learn about [Authentication for custom functions without a shared runtime](custom-functions-authentication.md).

## See also

- [Custom functions troubleshooting](custom-functions-troubleshooting.md)
- [Error handling for custom functions in Excel](custom-functions-errors.md)
- [Create custom functions in Excel](custom-functions-overview.md)
- [JavaScript-only runtime](../testing/runtimes.md#javascript-only-runtime)
