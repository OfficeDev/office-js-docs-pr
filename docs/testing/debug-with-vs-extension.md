---
title: Visual Studio Code Debugger Extension for Office Add-ins
description: 'Use a VS Code extension to debug your Office add-in.'
ms.date: 05/14/2020
localization_priority: Normal
---

# Visual Studio Code Debugger Extension for Office Add-ins

The Visual Studio Debugger Extension for Office Add-ins allows you to use F5 in VS Code to debug your Office Add-in against the Edge runtime. This debugging mode is dynamic, allowing you to set breakpoints while code is running. You can see changes in your code immediately while the debugger is attached, all without losing your debugging session. Your code changes also persist, so you can see the results of multiple changes to your code.

## Pre-Requisites

- [Visual Studio Code](https://code.visualstudio.com/) (must be run as an administrator)
- [Node.js (version 10+)](https://nodejs.org/)
- Windows 10
- [Microsoft Edge](https://www.microsoft.com/edge)
- [Yo Office](https://github.com/OfficeDev/generator-office)

These instructions assume you have experience using the command line, understand basic JavaScript, and have created an Office add-in project before using the Yo Office generator. If you haven't done this before, consider visiting one of our tutorials, like this [Excel Office Add-in tutorial](../tutorials/excel-tutorial.md).

## Install and use

1. [Create a new project with the latest version of Yo Office](https://docs.microsoft.com/office/dev/add-ins/quickstarts/excel-quickstart-jquery?tabs=yeomangenerator). Follow the prompts within the command line to set up your project. You can choose any language or type of project to suit your needs.

2. Open a command prompt as administrator
   ![Command prompt options, including "run as administrator" in Windows 10](../images/run-as-administrator-vs-code.jpg)

3. Navigate to your project directory.

4. Run the following command to open your project in Visual Studio Code as an administrator.

```command line
code .
```

Once Visual Studio Code is open, navigate manually to the project folder.

> [!TIP]
> To open Visual Studio Code as an administrator, select the **run as administrator** option when opening Visual Studio Code from the Start Menu in Windows.

5. Within VS Code, select **CTRL + SHIFT + X** to open the Extensions bar. Search for the "VS Code Debugger for Office Add-ins" extension and install it.

6. In the .vscode folder of your project, you will find a launch.json file. At the end of the file, add the following code to the configurations section of the file.

```JSON
{
  "type": "office-addin",
  "request": "attach",
  "name": "Attach to Office Add-ins",
  "port": 9222,
  "trace": "verbose",
  "url": "https://localhost:3000/taskpane.html?_host_Info=HOST$Win32$16.01$en-US$$$$0",
  "webRoot": "${workspaceFolder}",
  "timeout": 45000
}
```

7. In the section of JSON you just copied, find the "url" section. In this URL, you will need to replace the uppercase HOST text with the host application for your Office add-in. For example, if your Office add-in is for Excel, your URL value would be "https://localhost:3000/taskpane.html?_host_Info=<strong>Excel</strong>$Win32$16.01$en-US$\$\$\$0".

8. Open the command prompt and ensure you are at the root folder of your project. Run the command `npm start` to start the dev server. When your add-in loads in Excel, open the task pane.

9. Return to Visual Studio Code and choose **View > Debug** or enter **Ctrl+Shift+D** to switch to debug view.

10. From the Debug options, choose **Attach to Office Add-ins**. Select **F5** or choose **Debug -> Start Debugging** from the menu to begin debugging.

11. Set a breakpoint in your project's task pane file. You can set breakpoints in VS Code by hovering next to a line of code and selecting the red circle which appears.

![A red circle appears on a line of code in VS Code](../images/set-breakpoint.jpg)

12. Run your add-in. You will see that breakpoints have been hit and you can inspect local variables.
