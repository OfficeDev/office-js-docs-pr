---
title: Debug a function command with a non-shared runtime
description: Learn how to debug function commands.
ms.date: 06/17/2025
ms.localizationpriority: medium
---

# Debug a function command with a non-shared runtime

> [!IMPORTANT]
> If your add-in is [configured to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md), debug the code behind the function command just as you would the code behind a task pane. See [Debug Office Add-ins](debug-add-ins-overview.md) and note that a function command in an add-in with a [shared runtime](runtimes.md#shared-runtime) is *not* a special case as described in that article.

> [!NOTE]
> This article assumes that you're familiar with [function commands](../design/add-in-commands.md#types-of-add-in-commands).

Function commands don't have a UI, so a debugger can't be attached to the process in which the function runs on desktop Office. (Outlook add-ins being developed on Windows are an exception to this. See [Debug function commands in Outlook add-ins on Windows](#debug-function-commands-in-outlook-add-ins-on-windows) later in this article.) So function commands, in add-ins with a non-shared runtime, must be debugged in Office on the web where the function runs in the overall browser process. Use the following steps.

1. [Sideload the add-in in Office on the web](sideload-office-add-ins-for-testing.md), and then select the button or menu item that runs the function command. This is necessary to load the code file for the function command.
1. Open the browser's developer tools. This is usually done by pressing <kbd>F12</kbd>. The debugger in the tools attaches to the browser process.
1. Apply breakpoints to the code as needed for the function command.
1. Rerun the function command. The process stops on your breakpoints.

> [!TIP]
> For more detailed information, see [Debug add-ins in Office on the web](debug-add-ins-in-office-online.md).

## Debug function commands in Outlook add-ins on Windows

If your development computer is Windows, there's a way that you can debug a function command on Outlook desktop. See [Debug function commands in Outlook add-ins](../outlook/debug-ui-less.md).

## See also

- [Runtimes in Office Add-ins](runtimes.md)
