---
title: Group element in the manifest file
description: Defines a group of UI controls in a tab. 
ms.date: 01/29/2021
localization_priority: Normal
---

# Group element

Defines a group of UI controls in a tab. On custom tabs, the add-in can create multiple groups. Add-ins are limited to one custom tab.

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
|  [Control](#control)    | No |  Represents a Control object. Can be zero or more.  |
|  [OfficeControl](#officecontrol)  | No | Represents one of the built-in Office controls. Can be zero or more. |
|  [OverriddenByRibbonApi](overriddenbyribbonapi.md)      | No |  Specifies whether the group should appear on application and platform combinations that support custom contextual tabs.  |

### Label

Required. The label of the group. The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.

### Icon

Required. If a tab contains a lot of groups and the program window is resized, the specified image may display instead.

### Control

Optional, but if not present there must be at least one **OfficeControl**. For details about the types of controls that are supported, see the [Control](control.md) element. The order of **Control** and **OfficeControl** in the manifest is interchangeable and they can be intermingled if there are multiple elements, but all must be below the **Icon** element.

```xml
<Group id="contosoCustomTab.grp1">
    <Label resid="CustomTabGroupLabel"/>
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

### OfficeControl

Optional, but if not present there must be at least one **Control**. Include one or more built-in Office controls in the group with `<OfficeControl>` elements. The `id` attribute specifies the ID of the built-in Office control. To find the ID of a control, see [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups). The order of **Control** and **OfficeControl** in the manifest is interchangeable and they can be intermingled if there are multiple elements, but all must be below the **Icon** element.

```xml
<Group id="contosoCustomTab.grp1">
    <Label resid="CustomTabGroupLabel"/>
    <Icon>
        <bt:Image size="16" resid="blue-icon-16" />
        <bt:Image size="32" resid="blue-icon-32" />
        <bt:Image size="80" resid="blue-icon-80" />
    </Icon>
    <Control xsi:type="Button" id="Button2">
        <!-- information on the control -->
    </Control>
    <OfficeControl id="Superscript" />
    <!-- other controls, as needed -->
</Group>
```

### OverriddenByRibbonApi

Optional (boolean). Specifies whether the **Group** will be hidden on application and platform combinations that support an API that installs a custom contextual tab on the ribbon at runtime. The default value, if not present, is `false`. If used, **OverriddenByRibbonApi** must be the *first* child of **Group**. For more information, see [OverriddenByRibbonApi](overriddenbyribbonapi.md).

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="ContosoCustomTab.grp1">
      <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
      <!-- other child elements of the group -->
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```
