---
title: CustomTab element in the manifest file
description: On the ribbon, you specify which tab and group for their add-in commands.
ms.date: 02/25/2022
ms.localizationpriority: medium
---

# CustomTab element

Defines a custom tab for the Office ribbon. Add ribbon controls and groups for the add-in either to one of the build-in Office tabs or to your own custom tab. Use the **CustomTab** element to add a custom tab to the ribbon. On custom tabs, the add-in can have custom or built-in groups. Add-ins are limited to one custom tab.

> [!IMPORTANT]
> In Outlook on Mac, the **CustomTab** element is not available, but you can put *custom* groups of controls on one of the built-in [OfficeTab](officetab.md)s instead. You cannot put *built-in* groups on *built-in* tabs in Outlook on any platform.

**Add-in type:** Task pane, Mail

**Valid only in these VersionOverrides schemas**:

- Task pane 1.0
- Mail 1.0
- Mail 1.1

For more information, see [Version overrides in the manifest](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

> [!NOTE]
> Some child elements are not valid in the Mail schemas. See [Child elements](#child-elements).

**Associated with these requirement sets**:

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md)
- [AddinCommands 1.3](../requirement-sets/add-in-commands-requirement-sets.md). Required by some child elements. See [Child elements](#child-elements).

## Attributes

|  Attribute  |  Required  |  Description  |
|:-----|:-----|:-----|
|  [id](#id-attribute)  |  Yes  | A unique ID for the custom tab.|

### id attribute

Required. Unique identifier for the custom tab. It is a string with a maximum of 125 characters. This must be unique within the manifest.

## Child elements

|  Element |  Required  |  Description  |
|:-----|:-----|:-----|
|  [Group](group.md)      | No |  Defines a Group of commands.  |
|  [OfficeGroup](#officegroup)      | No |  Represents a built-in Office control group. **Important**: Not available in Outlook. |
|  [Label](#label-tab)      | Yes |  The label for the CustomTab.  |
|  [InsertAfter](#insertafter)      | No |  Specifies that the custom tab should be immediately after a specified built-in Office tab. **Important**: Only available in PowerPoint. |
|  [InsertBefore](#insertbefore)      | No |  Specifies that the custom tab should be immediately before a specified built-in Office tab. **Important**: Only available in PowerPoint. |

### Group

Optional, but if not present there must be at least one **OfficeGroup** element. See [Group element](group.md). The order of **Group** and **OfficeGroup** in the manifest should be the order you want them to appear on the custom tab. They can be intermingled if there are multiple elements, but all must be above the **Label** element.

### OfficeGroup

Optional, but if not present there must be at least one **Group** element. Represents a built-in Office control group. The **id** attribute specifies the ID of the built-in Office group. To find the ID of a built-in group, see [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups). The order of **Group** and **OfficeGroup** in the manifest should be the order you want them to appear on the custom tab. They can be intermingled if there are multiple elements, but all must be above the **Label** element.

> [!IMPORTANT]
> The **OfficeGroup** element is not available in Outlook. In PowerPoint, it is in preview for Mac and Windows; but is available for production add-ins in PowerPoint on the web.

**Add-in type:** Task pane

**Valid only in these VersionOverrides schemas**:

- Task pane 1.0

For more information, see [Version overrides in the manifest](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associated with these requirement sets**:

- [AddinCommands 1.3](../requirement-sets/add-in-commands-requirement-sets.md)

### Label (Tab)

Required. The label of the custom tab. The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.

**Add-in type:** Task pane, Mail

**Valid only in these VersionOverrides schemas**:

- Task pane 1.0
- Mail 1.0
- Mail 1.1

For more information, see [Version overrides in the manifest](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associated with these requirement sets**:

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md)

### InsertAfter

Optional. Specifies that the custom tab should be immediately after a specified built-in Office tab. The value of the element is the ID of the built-in tab, such as `TabHome` or `TabReview`.  For a list of built-in tabs, see [OfficeTab](officetab.md). If present, must be after the **Label** element. You cannot have both **InsertAfter** and **InsertBefore**.

> [!IMPORTANT]
> The **InsertAfter** element is only available in PowerPoint.

**Add-in type:** Task pane

**Valid only in these VersionOverrides schemas**:

- Task pane 1.0

For more information, see [Version overrides in the manifest](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associated with these requirement sets**:

- [AddinCommands 1.3](../requirement-sets/add-in-commands-requirement-sets.md)

### InsertBefore

Optional. Specifies that the custom tab should be immediately before a specified built-in Office tab. The value of the element is the ID of the built-in tab, such as `TabHome` or `TabReview`. The value of the element is the ID of the built-in tab, such as `TabHome` or `TabReview`.  For a list of built-in tabs, see [OfficeTab](officetab.md). If present, must be after the **Label** element. You cannot have both **InsertAfter** and **InsertBefore**.

> [!IMPORTANT]
> The **InsertBefore** element is only available in PowerPoint.

**Add-in type:** Task pane

**Valid only in these VersionOverrides schemas**:

- Task pane 1.0

For more information, see [Version overrides in the manifest](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associated with these requirement sets**:

- [AddinCommands 1.3](../requirement-sets/add-in-commands-requirement-sets.md)


## Examples

The following markup example adds the Office Paragraph control group to a custom tab and positions it to appear just after a custom group.

```xml
<ExtensionPoint xsi:type="ContosoRibbonTab">
  <CustomTab id="Contoso.TabCustom1">
    <Group id="Contoso.TabCustom1.group1">
       <!-- additional markup omitted -->
    </Group>
    <OfficeGroup id="Paragraph" />
    <Label resid="customTabLabel1" />
  </CustomTab>
</ExtensionPoint>
```

The following markup example adds the Office Superscript control to a custom group and positions it to appear just after a custom button.

```xml
<ExtensionPoint xsi:type="ContosoRibbonTab">
  <CustomTab id="Contoso.TabCustom2">
    <Group id="Contoso.TabCustom2.group2">
        <Label resid="residCustomTabGroupLabel"/>
        <Icon>
            <bt:Image size="16" resid="blue-icon-16" />
            <bt:Image size="32" resid="blue-icon-32" />
            <bt:Image size="80" resid="blue-icon-80" />
        </Icon>
        <Control xsi:type="Button" id="Contoso.Button2">
            <!-- information on the control omitted -->
        </Control>
        <OfficeControl id="Superscript" />
        <!-- other controls, as needed -->
    </Group>
    <Label resid="customTabLabel1" />
  </CustomTab>
</ExtensionPoint>
```
