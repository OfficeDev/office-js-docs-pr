---
title: Permissions element in the manifest file
description: ''
ms.date: 03/19/2019
localization_priority: Normal
---

# Permissions element

Specifies the level of API access for your Office Add-in; you should request permissions based on the principle of least privilege.

**Add-in type:** Content, Task pane, Mail

## Syntax

For content and task pane add-ins:

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

For mail add-ins

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## Contained in

[OfficeApp](officeapp.md)

## Remarks

For more detail, see [Requesting permissions for API use in add-ins](/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) and [Understanding Outlook add-in permissions](../outlook/understanding-outlook-add-in-permissions.md).
