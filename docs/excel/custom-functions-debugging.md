---
ms.date: 02/21/2019
description: Debug your custom functions in Excel.
title: Custom functions debugging (preview)
localization_priority: Normal
---
# Custom functions debugging (preview)

Debugging for custom functions can be accomplished by multiple means, depending on what platform you're using. Methods differ between [Office Online for Windows](#windows-10-and-microsoft-edge), [Office Online for Mac](#mac-and-chrome), Office Desktop for Windows, and Office Desktop for Mac. You can also issue debugging commands through [the command line directly](#using-the-command-line).

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## Office Online
### Windows 10 and Microsoft Edge

1. Open your custom functions root project folder in VS Code.
2. You may be prompted to install recommended extensions if you've based this project off of the latest yo office generator project. Select **install all**, which will install [Debugger for Microsoft Edge](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge) and [Debugger for Chrome](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-chrome) extensions for VS Code. If you're not prompted to install these, install these extensions to VS Code manually.
3. Select **Terminal | Run Task** and type or select Dev Server.
4. Switch to Debug View selecting **View | Debug** or by entering **Ctrl+Shift+D**.
5. From the Debug options, choose Office Online (Edge).
6. Open Excel Online using Edge and create a new workbook.
7. Copy the URL for this new workbook.
8. In VS Code, select **F5** to begin debugging. A prompt will appear, which asks for the URL of your document.
9. Paste in the URL for your workbook.
10. At this point, debugging is active. You can set breakpoints within VS Code.
11. To see your breakpoints applied, you'll go back to the Edge window with your open workbook and insert your add-in.
    - Select **Insert | Office Add-ins**. This will open a pop-up window.
    - Select **Manage My Add-ins | Upload Add-in** in the new pop-up window.
    - Browse for the manifest file of your add-in and select Upload.

### Mac and Chrome

1. Open your custom functions root project folder in VS Code.
2. You may be prompted to install recommended extensions if you've based this project off of the latest yo office generator project. Select **install all**, which will install [Debugger for Microsoft Edge](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge) and [Debugger for Chrome](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-chrome) extensions for VS Code. If you're not prompted to install these, install these extensions to VS Code manually.
3. Select **Terminal | Run Task** and type or select Dev Server.
4. Switch to Debug View selecting **View | Debug** or by entering **Ctrl+Shift+D**.
5. From the Debug options, choose Office Online (Chrome).
6. Open Excel Online using Chrome and create a new workbook.
7. Copy the URL for this new workbook.
8. In VS Code, select **F5** to begin debugging. A prompt will appear, which asks for the URL of your document.
9. Paste in the URL for your workbook.
10. At this point, debugging is active. You can set breakpoints within VS Code.
11. To see your breakpoints applied, you'll go back to the Chrome window with your open workbook and insert your add-in.
    - Select **Insert | Office Add-ins**. This will open a pop-up window.
    - Select **Manage My Add-ins | Upload Add-in** in the new pop-up window.
    - Browse for the manifest file of your add-in and select Upload.

## Office Desktop

### For Windows Or Mac

1. Open your custom functions root project folder in VS Code.
3. Select **Terminal | Run Task** and type or select Dev Server.
4. Switch to Debug View selecting **View | Debug** or by entering **Ctrl+Shift+D**.
5. From the Debug options, choose Office Desktop.
6. Select **F5** to begin debugging. A new Excel workbook will pop up, with your add-in already sideloaded and ready to use.
7. To debug, set breakpoints within VS Code.

## Using the command line

To debug using the command line, follow the same step sequences for platform and product, but replace VS Code terminal and debugging commands with statements below, depending on your needs:

### Watch and build your project

- `npm run watch`: builds for development and automatically rebuilds when a source file is saved

- `npm run build-dev`: builds for development once.

- `npm run build`: builds for production.

### Start the dev-server

- `npm run dev-server`: runs the web server used for development

### Debug

- If debugging for desktop, use `npm start desktop -- --app excel`. You can also use `npm stop desktop -- --app excel` to stop debugging.

- If debugging online, use `npm start web`. You'll need to manually open a new workbook using Excel Online and insert your add-in (for help, see the following note).  Select **F12** to use your browser's debugging tools.

> [!NOTE]
> To insert your add-in in Excel Online,  select **Insert | Office Add-ins**. This opens a pop-up window where you can select **Manage My Add-ins | Upload Add-in**. Browse for the manifest file of your add-in and select **Upload**.

## See also

* [Custom functions metadata](custom-functions-json.md)
* [Runtime for Excel custom functions](custom-functions-runtime.md)
* [Custom functions best practices](custom-functions-best-practices.md)
* [Custom functions changelog](custom-functions-changelog.md)
* [Excel custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md)
