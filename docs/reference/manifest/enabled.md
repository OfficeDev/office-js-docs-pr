---
title: Enabled element in the manifest file
description: 'Learn how to specify that an Add-in Command is disabled when the add-in launches.'
ms.date: 11/16/2021
ms.localizationpriority: medium
---

# Enabled element

Specifies whether a [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) control is enabled when the add-in launches. The **Enabled** element is a child element of [Control](control.md). If it is omitted, the default is `true`.

**Add-in type:** Task pane

**Valid only in these VersionOverrides schemas**:

- Task pane 1.0

For more information, see [Version overrides in the manifest](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associated with these requirement sets**:

- [RibbonApi 1.0](../requirement-sets/ribbon-api-requirement-sets.md)

This element is only valid in Excel; that is, when the `Name` attribute of the [Host](host.md) element is "Workbook".

The parent control can also be programmatically enabled and disabled. For more information, see [Enable and Disable Add-in Commands](../../design/disable-add-in-commands.md).

## Example

```xml
<Enabled>false</Enabled>
```
