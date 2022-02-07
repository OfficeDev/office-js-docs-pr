---
title: Items element in the manifest file
description: Specifies the items in a menu.
ms.date: 02/04/2022
ms.localizationpriority: medium
---

# Items element

Specifies the items in a menu.

**Add-in type:** Task pane, Mail

**Valid only in these VersionOverrides schemas**:

- Task pane 1.0
- Mail 1.0
- Mail 1.1

For more information, see [Version overrides in the manifest](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associated with these requirement sets**:

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md) when the parent **VersionOverrides** is type Taskpane 1.0.
- [Mailbox 1.3](../../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md) when the parent **VersionOverrides** is type Mail 1.0.
- [Mailbox 1.5](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) when the parent **VersionOverrides** is type Mail 1.1.

## Syntax

```XML
<Items>
...  
</Items>  
```

## Contained in

[Control element of type Menu](control-menu.md)

## Must contain

[Item](item.md)

## Examples

For examples, see [Control of type Menu](control-menu.md).