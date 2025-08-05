---
title: Automatically open a task pane when an add-in is installed
description: Learn how to configure an Office Add-in to open automatically when it's installed.
ms.topic: how-to
ms.date: 02/12/2025
ms.localizationpriority: medium
---


# Automatically open a task pane when an add-in is installed

You can configure your add-in's task pane to launch immediately after it's installed. This feature increases usage. 

By default, task pane add-ins that do *not* include any [add-in commands](../design/add-in-commands.md) open the task pane immediately upon installation. However, when an add-in has one or more add-in commands, then the user is notified of new add-in, but the add-in doesn't launch automatically. This historic default behavior is changing so add-ins that do have add-in commands will launch automatically in some situations. In addition, if the add-in has more than one task pane page, it's possible for you to control whether the add-in launches upon installation and, if so, which page opens in the task pane.

> [!NOTE]
> 
> - This feature applies only to add-ins installed by an end-user, not to centrally deployed add-ins.
> - This feature doesn't apply to Content add-ins or Mail (Outlook) add-ins.
> - This feature applies only to add-ins that have at least one add-in command of [the type "task pane command"](../design/add-in-commands.md#types-of-add-in-commands).

## New behavior

The new behavior is as follows:

- If the add-in has just one [task pane command](../design/add-in-commands.md#types-of-add-in-commands), then the add-in's ribbon tab is selected and the task pane opens automatically upon installation. You don't need to configure anything.
- If the add-in has multiple task pane commands, and one is configured to be the default (see [Configure default task pane](#configure-default-task-pane-in-the-manifest)), then the add-in's ribbon tab is selected and the default task pane opens automatically upon installation.
- If the add-in has multiple task pane commands, but none is configured to be the default, then the add-in's ribbon tab is selected automatically upon installation and a callout appears near it notifying the user of the new add-in, but no task pane is opened. This is the same as the historic default behavior.

> [!NOTE]
> If for any reason, the add-in command that launches the task pane cannot be manually selected by a user at start up, such as when it's [configured to be disabled](../design/disable-add-in-commands.md) at start up, then it won't be automatically opened regardless of configuration. 

## Configure default task pane in the manifest

The process for specifying the default task pane depends on the type of manifest the add-in uses.

# [Unified manifest for Microsoft 365](#tab/jsonmanifest)

[!include[Unified manifest host application support note](../includes/unified-manifest-support-note.md)]

To specify the default task pane, find the runtime object in the [`"runtimes"`](/microsoft-365/extensibility/schema/element-extensions#runtimes) array whose [`"code.page"`](/microsoft-365/extensibility/schema/extension-runtime-code#page) property is set to the URL of the page that you want to be the default. Ensure that the [`"actions"`](/microsoft-365/extensibility/schema/extension-runtimes-actions-item) array in this same runtime object has at least one action whose `"type"` value is `"openPage"`. Add a `"view"` property to this action object and set it to `"Office.AutoShowTaskpaneWithDocument"`. You can only set this value on one action object and it must be an action of type `"openPage"`. If you set this value on multiple actions, the first occurrence of the value will be recognized as the default and the others will be ignored.

The following example shows a `"view"` value set to `"Office.AutoShowTaskpaneWithDocument"`.

```json
"runtimes": [
    {
        ...
        "code": {
            "page": "https://contoso.com/taskpane.html"
        },
        "actions": [
            {
                "id": "ShowTaskPane",
                "type": "openPage",
                "view": "Office.AutoShowTaskpaneWithDocument"
            }
        ]
    }
]
```

# [Add-in only manifest](#tab/xmlmanifest)

To designate a task pane as the default, add a [TaskpaneId](/javascript/api/manifest/action#taskpaneid) element as the first child of the `<Action>` element and set its value to **Office.AutoShowTaskpaneWithDocument**. The following is an example.

```xml
<Action xsi:type="ShowTaskpane">
    <TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>
    <SourceLocation resid="Contoso.Taskpane.Url" />
</Action>
```

---

> [!TIP]
> If you want your add-in to automatically launch whenever the user reopens the document, you need to take further configuration steps. For details and advice about when to use this feature, see [Automatically open a task pane with a document](automatically-open-a-task-pane-with-a-document.md). 

## See also

- [Automatically open a task pane with a document](automatically-open-a-task-pane-with-a-document.md)
