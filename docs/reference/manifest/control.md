---
title: Control element in the manifest file
description: Defines a control that executes an action or launches a task pane.
ms.date: 02/04/2022
ms.localizationpriority: medium
---

# Control element

Defines a control that executes an action or launches a task pane. A **Control** element can be either a button or a menu option. At least one **Control** must be included in a [Group](group.md) element.

**Add-in type:** Task pane, Mail

**Valid only in these VersionOverrides schemas**:

- Task pane 1.0
- Mail 1.0
- Mail 1.1

For more information, see [Version overrides in the manifest](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associated with these requirement sets**:

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md) (For a task pane add-in.)
- Some child elements may be associated with additional requirement sets.

## Attributes

|  Attribute  |  Required  |  Description  |
|:-----|:-----|:-----|
|**xsi:type**|Yes|The type of control being defined. Can be `Button`, `Menu`, or `MobileButton`. |
|**id**|Yes|The ID of the control element. Can be a maximum of 125 characters. Must be unique across all **Control** elements in the manifest.|

> [!NOTE]
> The `MobileButton` value for **xsi:type** is defined in VersionOverrides schema 1.1. It only applies to the **Control** elements contained within a [MobileFormFactor](mobileformfactor.md) element.

## Child elements

The valid child elements depend on the value of the **xsi:type** attribute.

- [Button type of Control element](control-button.md)
- [Menu type of Control element](control-menu.md)
- [MobileButton type of Control element](control-mobilebutton.md)
