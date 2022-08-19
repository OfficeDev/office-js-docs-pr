---
title: Debug the initialize and onReady functions
description: Learn how to debug the Office.initialize and Office.onReady functions.
ms.date: 07/11/2022
ms.localizationpriority: medium
---

# Debug the initialize and onReady functions

> [!NOTE]
> This article assumes that you are familiar with [Initialize your Office Add-in](../develop/initialize-add-in.md).

The paradox of debugging the [Office.initialize](/javascript/api/office#office-office-initialize-function(1)) and [Office.onReady](/javascript/api/office#office-office-onready-function(1)) functions is that a debugger can only attach to a process that is running, but these functions run immediately as the add-in's Configure your Office Add-in to use a shared runtime process starts up, before a debugger can attach. In most situations, restarting the add-in after a debugger is attached doesn't help because restarting the add-in closes the original Configure your Office Add-in to use a shared runtime process *and the attached debugger* and starts a new process that has no debugger attached.

Fortunately, there is an exception. You can debug these functions using Office on the web, with the following steps.

1. Sideload and run the add-in in Office on the web. This is usually done by opening an add-in's task pane or running a [function command](../design/add-in-commands.md#types-of-add-in-commands). *The add-in runs in the overall browser process, not a separate process as it would in desktop Office.*
1. Open the browser's developer tools. This is usually done by pressing F12. The debugger in the tools attaches to the browser process.
1. Apply breakpoints as needed to the code in the `Office.initialize` or `Office.onReady` function.
1. *Relaunch the add-in's task pane or the function command* just as you did in step 1. This action does *not* close the browser process or the debugger. The `Office.initialize` or `Office.onReady` function runs again and processing stops on your breakpoints.

> [!TIP]
> For more detailed information, see [Debug add-ins in Office on the web](debug-add-ins-in-office-online.md).

## See also

- [Configure your Office Add-in to use a shared runtimes in Office Add-ins](Configure your Office Add-in to use a shared runtimes.md)
