---
title: Permissions element in the manifest file
description: The Permissions element specifies the API access level for your Office Add-in.
ms.date: 06/26/2020
ms.localizationpriority: medium
---

# Permissions element

Specifies the level of API access for your Office Add-in; you should request permissions based on the principle of least privilege.

**Add-in type:** Content, Task pane, Mail

## Syntax

For content and task pane add-ins:

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

For mail add-ins:

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## Contained in

[OfficeApp](officeapp.md)

## Remarks

For more details, see [Requesting permissions for API use in content and task pane add-ins](../../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md) and [Understanding Outlook add-in permissions](../../outlook/understanding-outlook-add-in-permissions.md).
