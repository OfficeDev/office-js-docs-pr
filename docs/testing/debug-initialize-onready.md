---
title: Debug the initialize and onReady functions
description: Learn how to debug the Office.initialize and Office.onReady functions.
ms.date: 11/06/2025
ms.localizationpriority: medium
---

# Debug the initialize and onReady functions

> [!NOTE]
> This article assumes that you're familiar with [Initialize your Office Add-in](../develop/initialize-add-in.md).

The paradox of debugging the [Office.initialize](/javascript/api/office#office-office-initialize-function(1)) and [Office.onReady](/javascript/api/office#office-office-onready-function(1)) functions is that a debugger can only attach to a process that's running, but these functions run immediately as the add-in's runtime process starts up, before a debugger can attach. In most situations, restarting the add-in after a debugger is attached doesn't help because restarting the add-in closes the original runtime process *and the attached debugger* and starts a new process that has no debugger attached.

Fortunately, there are two ways that you can debug these functions that are described in the following sections.

## Debug using Office on the web

To debug with Office on the web, use the following steps.

1. Sideload and run the add-in in Office on the web. This is usually done by opening an add-in's task pane or running a [function command](../design/add-in-commands.md#types-of-add-in-commands). *The add-in runs in the overall browser process, not a separate process as it would in desktop Office.*
1. Open the browser's developer tools. This is usually done by pressing <kbd>F12</kbd>. The debugger in the tools attaches to the browser process.
1. Apply breakpoints as needed to the code in the `Office.initialize` or `Office.onReady` function.
1. *Relaunch the add-in's task pane or the function command* just as you did in step 1. This action does *not* close the browser process or the debugger. The `Office.initialize` or `Office.onReady` function runs again and processing stops on your breakpoints.

> [!TIP]
> For more detailed information, see [Debug add-ins in Office on the web](debug-add-ins-in-office-online.md).

## Debug using Office on Windows

[!INCLUDE[Automatically open the Microsoft Edge developer tools to debug initialization](../includes/auto-open-webview2-dev-tools.md)]

## See also

- [Runtimes in Office Add-ins](runtimes.md)
