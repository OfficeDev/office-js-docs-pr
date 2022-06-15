---
title: Debug Office Add-ins
description: Find the Office Add-in debugging guidance for your development environment.
ms.date: 06/15/2022
ms.localizationpriority: high
---

# Overview of debugging Office Add-ins

Debugging Office Add-ins is essentially the same as debugging any web application. However, a single set of tools won't work for all add-in developers. This is because add-ins can be developed on different operating systems and run cross-platform. This article helps you find the detailed debugging guidance for your development environment.

> [!TIP]
> This article is concerned with debugging in the narrow sense of setting breakpoints and stepping through code. For guidance on testing and troubleshooting, start with [Test Office Add-ins](test-debug-office-add-ins.md) and [Troubleshoot development errors with Office Add-ins](troubleshoot-development-errors.md).

> [!NOTE]
> Although you should *test* your add-in on all the platforms that you want to support, you'll only very rarely need to *debug* on an environment different from your development computer. For this reason, this article uses "your development computer" and "your development environment" to refer to the environment on which you're debugging. If a problem in the code occurs only on a platform other than the one on your development computer, and you need to set breakpoints or step through code to solve it, then the environment on which you're debugging isn't literally your development environment.

## Server-side or client-side?

Debugging the server-side code of an Office Add-in is the same as debugging the server-side of any web application. See the debugging instructions for your IDE or other tools. The following are examples for some of the most popular tools.

- [Debug ASP.NET or ASP.NET Core apps in Visual Studio](/visualstudio/debugger/how-to-enable-debugging-for-aspnet-applications)
- [Debugging Express](https://expressjs.com/en/guide/debugging.html)
- [Node.js Debugging Guide](https://nodejs.org/en/docs/guides/debugging-getting-started/)
- [Node.js debugging in VS Code](https://code.visualstudio.com/docs/nodejs/nodejs-debugging)
- [Webpack Debugging](https://webpack.js.org/contribute/debugging/)

The rest of this article is concerned only with debugging client-side JavaScript (which may be transpiled from TypeScript).

To find guidance for debugging client-side code, the first variable is the operating system of your development computer.

- [Windows](#debug-on-windows)
- [Mac](#debug-on-mac)
- [Linux or other Unix variant](#debug-on-linux)

## Debug on Windows

The following provides general guidance to debugging on Windows. There are special instructions for debugging custom functions in Excel and event-based add-ins in Outlook. See [Special cases in Windows](#special-cases-in-windows) later in this section. Debugging on Windows depends on your IDE:

- **Visual Studio**: Debug using the browser's F12 tools. See [Debug Office Add-ins in Visual Studio](../develop/debug-office-add-ins-in-visual-studio.md).
- **Visual Studio Code**: Debug using the [Add-in Debugger Extension for Visual Studio Code](debug-with-vs-extension.md).
- **Any other IDE** (or you don't want to debug inside your IDE): Use the developer tools that are associated with the browser runtime that add-ins use on your development computer. See one of the following:

    - [Debug add-ins using developer tools for Internet Explorer](debug-add-ins-using-f12-tools-ie.md)
    - [Debug add-ins using developer tools for Edge Legacy](debug-add-ins-using-devtools-edge-legacy.md)
    - [Debug add-ins using developer tools in Microsoft Edge (Chromium-based)](debug-add-ins-using-devtools-edge-chromium.md)

For information about which browser runtime is being used, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).

> [!TIP]
> [!INCLUDE[Identify the webview through the add-in UI](../includes/identify-webview-in-ui.md)]

### Special cases in Windows

To debug custom functions on Windows, see [Custom functions debugging](../excel/custom-functions-debugging.md).

To debug event-based add-ins in Outlook, see [Debug your event-based Outlook add-in](../outlook/debug-autolaunch.md). The process requires Visual Studio Code.

## Debug on Mac

The following provides general guidance to debugging on Mac. There are special instructions for debugging custom functions in Excel. See [Special cases in Mac](#special-cases-in-mac) later in this section.

- If you're using Visual Studio Code, debug using the [Add-in Debugger Extension for Visual Studio Code](debug-with-vs-extension.md).
- For any other IDE, use the Safari Web Inspector. Instructions are in [Debug Office Add-ins on a Mac](debug-office-add-ins-on-ipad-and-mac.md).

### Special cases in Mac

To debug custom functions on Mac, see [Custom functions debugging](../excel/custom-functions-debugging.md).

## Debug on Linux

There is no desktop version of Office for Linux, so you'll need to [sideload the add-in to Office on the web](sideload-office-add-ins-for-testing.md) to test and debug it. Debugging guidance is in [Debug add-ins in Office on the web](debug-add-ins-in-office-online.md).

> [!NOTE]
> We don't recommend that you develop Office Add-ins on a Linux computer except in the unusual case where you can be sure that all the add-in's users will be accessing the add-in through Office on the web from a Linux computer.

## Debug add-ins in staging or production

To debug an add-in that is already in staging or production, attach a debugger from the UI of the add-in. For instructions, see [Attach a debugger from the task pane](attach-debugger-from-task-pane.md).