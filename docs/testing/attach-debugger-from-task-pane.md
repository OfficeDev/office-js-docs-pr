---
title: Attach a debugger from the task pane
description: ''
ms.date: 10/17/2018
---

# Attach a debugger from the task pane

In Office 2016 for Windows, Build 77xx.xxxx or later, you can attach the debugger from the task pane. The attach debugger feature will directly attach the debugger to the correct Internet Explorer process for you. You can attach a debugger regardless of whether you are using Yeoman Generator, Visual Studio Code, node.js, Angular, or another tool. 

To launch the **Attach Debugger** tool, choose the top right corner of the task pane to activate the **Personality** menu (as shown in the red circle in the following image).   

> [!NOTE]
> - Currently the only supported debugger tool is [Visual Studio 2015](https://www.visualstudio.com/downloads/) with [Update 3](https://msdn.microsoft.com/library/mt752379.aspx) or later. If you don't have Visual Studio installed, selecting the **Attach Debugger** option doesn’t result in any action.   
> - You can only debug client-side JavaScript with the **Attach Debugger** tool. To debug server-side code, such as with a Node.js server, you have many options. For information on how to debug with Visual Studio Code, see [Node.js Debugging in VS Code](https://code.visualstudio.com/docs/nodejs/nodejs-debugging). If you are not using Visual Studio Code, search for "debug Node.js" or "debug {name-of-server}".

![Screenshot of Attach Debugger menu](../images/attach-debugger.png)

Select **Attach Debugger**. This launches the **Visual Studio Just-in-Time Debugger** dialog box, as shown in the following image. 

![Screenshot of Visual Studio JIT Debugger dialog](../images/visual-studio-debugger.png)

In Visual Studio, you will see the code files in **Solution Explorer**.   You can set breakpoints to the line of code you want to debug in Visual Studio.

If you don't see the Personality menu, you can debug your add-in using Visual Studio. Ensure your task pane add-in is open in Office, and then follow these steps:

1. In Visual Studio, choose **DEBUG** > **Attach to Process**.
2. In **Attach to Process**, choose all of the available Iexplore.exe processes, and then choose the **Attach** button.

For more information about debugging in Visual Studio, see the following:

-	To launch and use the DOM Explorer in Visual Studio, see Tip 4 in the [Tips and Tricks](https://blogs.msdn.microsoft.com/officeapps/2013/04/16/building-great-looking-apps-for-office-using-the-new-project-templates/#tips_tricks) section of the [Building great-looking apps for Office using the new project templates](https://blogs.msdn.microsoft.com/officeapps/2013/04/16/building-great-looking-apps-for-office-using-the-new-project-templates) blog post.
-	To set breakpoints, see [Using Breakpoints](https://docs.microsoft.com/visualstudio/debugger/using-breakpoints?view=vs-2015).
-	To use F12, see [Using the F12 developer tools](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85)).

## See also

- [Create and debug Office Add-ins in Visual Studio](../develop/create-and-debug-office-add-ins-in-visual-studio.md)
- [Publish your Office Add-in](../publish/publish.md)
