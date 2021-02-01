---
title: OverriddenByRibbonApi element in the manifest file
description: Learn how to specify that a custom tab, group, control, or menu item shouldn't appear when it is also part of a custom contextual tab.
ms.date: 01/29/2021
localization_priority: Normal
---

# OverriddenByRibbonApi element

Specifies whether a [CustomTab](customtab.md), [Group](group.md), [Button](control.md#button-control) control, [Menu](control.md#menu-dropdown-button-controls) control, or menu item will be hidden on host and platform combinations that support the API ([Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-)) that installs custom contextual tabs on the ribbon.

If it is omitted, the default is `false`. If it is used, it must be the *first* child element of its parent element.

> [!NOTE]
> For a full understanding of this element, please read [Create custom contextual tabs in Office Add-ins](../../design/contextual-tabs.md).

The purpose of this element is to create a fallback experience in an add-in that implements custom contextual tabs when the add-in is running on a host or platform that doesn't support custom contextual tabs. The essential strategy is that you define in the manifest one or more custom core tabs (that is, *noncontextual* custom tabs) that duplicate the ribbon customizations of the custom contextual tabs in your add-in. But you add `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` as the first child element of the **CustomTab**. The effect of doing so is the following:

- If the add-in runs on a host and platform that supports custom contextual tabs, then the custom tab won't appear on the ribbon. Instead, the custom contextual tab will be installed when the add-in calls the `requestCreateControls` method.
- If the add-in runs on a host or platform that *doesn't* support custom contextual tabs, then the custom core tab will appear on the ribbon.

There are more complex strategies for using this element. For details, see [Implement an alternate UI experience when custom contextual tabs aren't supported](../../design/contextual-tabs.md#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported).

## Example

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="TabCustom1">
    <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
    <Group id="ContosoCustomTab.grp1">
      <Control  xsi:type="Button" id="MyButton">
        <!-- child elements omitted -->
      </Control>
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```
