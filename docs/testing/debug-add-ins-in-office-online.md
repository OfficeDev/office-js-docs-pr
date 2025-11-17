---
title: Debug add-ins in Office on the web
description: How to use Office on the web to test and debug your add-ins.
ms.date: 12/19/2023
ms.localizationpriority: medium
---

# Debug add-ins in Office on the web

This article describes how to use Office on the web to debug your add-ins. Use this technique:

- To debug add-ins on a computer that isn't running Windows or the Office desktop client&mdash;for example, if you're developing on a Mac or Linux.
- As an alternative debugging process if you can't, or don't wish to, debug in an IDE, such as Visual Studio or Visual Studio Code.

This article assumes that you have an add-in project that needs to be debugged. If you just want to practice debugging on the web, create a new project using one of the quick starts for specific Office applications, such as this [quick start for Word](../quickstarts/word-quickstart-yo.md).

## Debug your add-in

To debug your add-in by using Office on the web:

1. Run the project on localhost and sideload it to a document in Office on the web. For detailed sideloading instructions, see [Manually sideload Office Add-ins on the web](sideload-office-add-ins-for-testing.md#manually-sideload-an-add-in-to-office-on-the-web).

1. Open the browser's developer tools. This is usually done by pressing <kbd>F12</kbd>. Open the debugger tool and use it to set breakpoints and watch variables. For detailed help in using your browser's tool, see one of the following:

    - [Firefox](https://firefox-source-docs.mozilla.org/devtools-user/index.html)
    - [Safari](https://support.apple.com/guide/safari/use-the-developer-tools-in-the-develop-menu-sfri20948/mac)
    - [Debug add-ins using developer tools in Microsoft Edge](debug-add-ins-using-devtools-edge-chromium.md)

    > [!NOTE]
    > The new Outlook on Windows desktop client doesn't support the context menu or the keyboard shortcut to access the Microsoft Edge developer tools. Instead, you must run `olk.exe --devtools` from a command prompt. For more information, see the "Debug your add-in" section of [Develop Outlook add-ins for the new Outlook on Windows](../outlook/one-outlook.md#debug-your-add-in).

## Potential issues

The following are some issues that you might encounter as you debug.

- Some JavaScript errors that you see might originate from Office on the web.
- The browser might show an invalid certificate error that you'll need to bypass. The process for doing this varies with the browser and the various browsers' UIs for doing this change periodically. You should search the browser's help or search online for instructions. (For example, search for "Microsoft Edge invalid certificate warning".) Most browsers will have a link on the warning page that enables you to click through to the add-in page. For example, Microsoft Edge has a "Go on to the webpage (Not recommended)" link. But you'll usually have to go through this link every time the add-in reloads. For a longer lasting bypass, see the help as suggested.
- If you set breakpoints in your code, Office on the web might throw an error indicating that it's unable to save.

## See also

- [Best practices for developing Office Add-ins](../concepts/add-in-development-best-practices.md)
- [Troubleshoot user errors with Office Add-ins](testing-and-troubleshooting.md)
