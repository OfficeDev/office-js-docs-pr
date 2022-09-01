---
title: Automatically open a task pane when an add-in is installed
description: Learn how to configure an Office Add-in to open automatically when it's installed.
ms.date: 09/01/2022
ms.localizationpriority: medium
---


# Automatically open a task pane when an add-in is installed

You can configure your add-in's task pane to launch immediately after it's installed. This feature increases usage. 

By default, when a task pane add-in that does *not* include any [add-in commands](../design/add-in-commands.md) is installed, it's task pane opens immediately upon installation. But when the add-in does have one or more add-in commands, the add-in's ribbon tab is selected and a callout appears near it notifying the user of the new add-in, but the add-in doesn't launch until the user takes action to start it, such as selecting a button to open a task pane. This historic default behavior is changing so that automatic launching upon installation is restored in some situations to add-ins that include add-in commands. In addition, if the add-in has more than one task pane page, it's possible for you to control whether the add-in launches upon installation and, if so, which page opens in the task pane. 

> [!NOTE]
> 
> - This feature is currently available only in Office on the web. We are working to bring this behavior to other platforms, but currently they still exhibit the historic default behavior described earlier.
> - This feature applies only to add-ins installed by an end-user, not to centrally deployed add-ins.
> - This feature doesn't apply to Content add-ins.
> - This feature applies only to add-ins that have at least one add-in command of [the type "task pane command"](../design/add-in-commands.md#types-of-add-in-commands).

## New behavior

The new behavior is as follows:

- If the add-in has just one [task pane command](../design/add-in-commands.md#types-of-add-in-commands), then the add-in's ribbon tab is selected and the task pane opens automatically upon installation. You don't need to configure anything.
- If the add-in has multiple task pane commands, and one is configured to be the default (see [Configure default task pane](#configure-default-task-pane)), then the add-in's ribbon tab is selected and the default task pane opens automatically upon installation.
- If the add-in has multiple task pane commands, but none is configured to be the default, then the add-in's ribbon tab is selected automatically upon installation and a callout appears near it notifying the user of the new add-in, but no task pane is opened. This is the same as the historic default behavior.

> [!NOTE]
> If for any reason, the add-in command that launches the task pane cannot be manually selected by a user at start up, such as when it's [configured to be disabled](../design/disable-add-in-commands.md) at start up, then it won't be automatically opened regardless of configuration. 

## Configure default task pane

To designate a task pane as the default, add a [TaskpaneId](/javascript/api/manifest/action#taskpaneid) element as the first child of the **\<Action\>** element and set its value to **Office.AutoShowTaskpaneWithDocument**. The following is an example.

```xml
<Action xsi:type="ShowTaskpane">
    <TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>
    <SourceLocation resid="Contoso.Taskpane.Url" />
</Action>
```

> [!TIP]
> If you want your add-in to automatically launch whenever the user reopens the document, you need to take further configuration steps. For details and advice about when to use this feature, see [Automatically open a task pane with a document](automatically-open-a-task-pane-with-a-document.md). 

