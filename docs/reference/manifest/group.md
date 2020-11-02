---
title: Group element in the manifest file
description: Defines a group of UI controls in a tab. 
ms.date: 12/02/2019
localization_priority: Normal
---

# Group element

Defines a group of UI controls in a tab.  On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.

## Attributes

|  Attribute  |  Required  |  Description  |
|:-----|:-----|:-----|
|  [id](#id-attribute)  |  Yes  | A unique ID for the group.|

### id attribute

Required. Unique identifier for the group. It is a string with a maximum of 125 characters. This must be unique within the manifest or the group will fail to render.

## Child elements
|  Element |  Required  |  Description  |
|:-----|:-----|:-----|
|  [Label](#label)      | Yes |  The label for the CustomTab or a group.  |
|  [Icon](icon.md)      | Yes |  The image for a group.  |
|  [Control](#control)    | Yes |  Collection of one or more Control objects.  |

### Label 

Required. The label of the group. The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.

### Icon

Required. If a tab contains a lot of groups and the program window is resized, the specified image may display instead.

### Control
A group requires at least one control. For details about the types of controls that are supported, see the [Control](control.md) element.

```xml
<Group id="msgreadCustomTab.grp1">
    <Label resid="residCustomTabGroupLabel"/>
    <Icon>
        <bt:Image size="16" resid="blue-icon-16" />
        <bt:Image size="32" resid="blue-icon-32" />
        <bt:Image size="80" resid="blue-icon-80" />
    </Icon>
    <Control xsi:type="Button" id="Button2">
        <!-- information on the control -->
    </Control>
    <!-- other controls, as needed -->
</Group>
```
