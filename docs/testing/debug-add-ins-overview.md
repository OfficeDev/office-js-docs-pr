---
title: Debug Office Add-ins
description: Find the Office Add-in debugging guidance for your development environment.
ms.topic: overview
ms.date: 03/20/2025
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
- [Node.js Debugging Guide](https://nodejs.org/en/learn/getting-started/debugging)
- [Node.js debugging in VS Code](https://code.visualstudio.com/docs/nodejs/nodejs-debugging)
- [Webpack Debugging](https://webpack.js.org/contribute/debugging/)

The rest of this article is concerned only with debugging client-side JavaScript (which may be transpiled from TypeScript).

## Special cases

There are some special cases in which the debugging process differs from normal for a given combination of platform, Office application, and development environment. If you're debugging any of these special cases, use the links in this section to find the proper guidance. Otherwise, continue to [General guidance](#general-guidance).

- **Debugging the `Office.initialize` or `Office.onReady` function**: [Debug the initialize and onReady functions](debug-initialize-onready.md).
- **Debugging an Excel custom function in a *non-shared* runtime**: [Custom functions debugging in a non-shared runtime](../excel/custom-functions-debugging.md).
- **Debugging a [function command](../design/add-in-commands.md#types-of-add-in-commands) in a *non-shared* runtime**:
  - Outlook add-ins on a Windows development computer: [Debug function commands in Outlook add-ins](../outlook/debug-ui-less.md)
  - Other Office application add-ins or Outlook on a Mac development computer: [Debug a function command with a non-shared runtime](debug-function-command.md).
- **Debugging an event-based or spam-reporting add-in**: [Debug event-based and spam-reporting add-ins](debug-autolaunch.md).
- **Debugging an add-in in the new Outlook on Windows desktop client**: See the "Debug your add-in" section of [Develop Outlook add-ins for the new Outlook on Windows](../outlook/one-outlook.md#debug-your-add-in).
- **Debugging a Blazor-based add-in**: Debug the add-in the same way you would debug a Blazor web application. See [Debug ASP.NET Core Blazor WebAssembly](/aspnet/core/blazor/debug/).

## General guidance

To find guidance for debugging client-side code, the first variable is the operating system of your development computer.

- [Windows](#debug-on-windows)
- [Mac](#debug-on-mac)
- [Linux or other Unix variant](#debug-on-linux)

### Debug on Windows

The following provides general guidance to debugging on Windows. Debugging on Windows depends on your IDE.

- **Visual Studio**: Debug using the browser's F12 tools. See [Debug Office Add-ins in Visual Studio](../develop/debug-office-add-ins-in-visual-studio.md).
- **Any other IDE** (or you don't want to debug inside your IDE): Use the developer tools that are associated with the webview control that add-ins use on your development computer. See one of the following:

  - For the Trident webview: [Debug add-ins using developer tools for Internet Explorer](debug-add-ins-using-f12-tools-ie.md)
  - For the EdgeHTML webview: [Debug add-ins using developer tools for Edge Legacy](debug-add-ins-using-devtools-edge-legacy.md)
  - For the WebView2 webview: [Debug add-ins using developer tools in Microsoft Edge (Chromium-based)](debug-add-ins-using-devtools-edge-chromium.md)

For information about which runtime is being used, see [Browsers and webview controls used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md) and [Runtimes in Office Add-ins](runtimes.md).

> [!TIP]
> [!INCLUDE[Identify the webview through the add-in UI](../includes/identify-webview-in-ui.md)]

### Debug on Mac

Use the Safari Web Inspector. Instructions are in [Debug Office Add-ins on a Mac](debug-office-add-ins-on-ipad-and-mac.md).

### Debug on Linux

There is no desktop version of Office for Linux, so you'll need to [sideload the add-in to Office on the web](sideload-office-add-ins-for-testing.md) to test and debug it. Debugging guidance is in [Debug add-ins in Office on the web](debug-add-ins-in-office-online.md).

> [!NOTE]
> We don't recommend that you develop Office Add-ins on a Linux computer except in the unusual case where you can be sure that all the add-in's users will be accessing the add-in through Office on the web from a Linux computer.

## Debug add-ins in staging or production

To debug an add-in that is already in staging or production, attach a debugger from the UI of the add-in. For instructions, see [Attach a debugger from the task pane](attach-debugger-from-task-pane.md).

## Versions of office.js for debugging

There are debug versions of the Office JavaScript libraries. These versions are more human readable and easier to step through with a debugger. Use them when the Office JavaScript APIs aren't working as expected. Avoid using them when you publish and deploy your add-in.

The debug versions are found at the following CDN locations.

- Office JavaScript API library: `https://appsforoffice.microsoft.com/lib/1/hosted/office.debug.js`
- Office JavaScript API (preview) library: `https://appsforoffice.microsoft.com/lib/beta/hosted/office.debug.js`

## See also

- [Runtimes in Office Add-ins](runtimes.md)
