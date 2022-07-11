---
title: Debug a function command with a non-shared runtime
description: Learn how to debug function commands.
ms.date: 07/11/2022
ms.localizationpriority: medium
---

# Debug a function command with a non-shared runtime

> [!IMPORTANT]
> If your add-in is [configured to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md), you debug the code behind the function command just as you would the code behind a task pane. See [Debug Office Add-ins](debug-add-ins-overview.md) and note that a function command in an add-in with a shared runtime is *not* a special case as described in that article. 

> [!NOTE]
> This article assumes that you are familiar with [function commands]((../design/add-in-commands.md#types-of-add-in-commands)).

Function commands do not have a UI, so a debugger cannot be attached to the process in which the function runs on desktop Office. (Outlook add-ins being developed on Windows are an exception to this. See [Debugging function commands in Outlook add-ins](#debugging-function-commands-in-Outlook-add-ins) later in this article.) So function commands, in add-ins with a non-shared runtime, must be debugged on Office on the web where the function runs in the overall browser process. Use the following steps.

1. Sideload the add-in in Office on the web, and then select the button or menu item that runs the function command. This is necessary to load the code file for the function command. 
2. Open the browser's developer tools. This is usually done by pressing F12. The debugger in the tools attaches to the browser process.
3. Apply breakpoints as needed to the code for the function command.
4. Rerun the function command. The process stops on your breakpoints. 

> [!TIP]
> There is more detailed information in [Debug add-ins in Office on the web](debug-add-ins-in-office-online.md).

## Debugging function commands in Outlook add-ins on Windows

If your development computer is Windows, there is a way that you can debug a function command on Outlook desktop. See [Debug function commands in Outlook add-ins](../outlook/debug-ui-less.md).