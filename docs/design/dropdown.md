---
title: DropDown component in Office UI Fabric
description: ''
ms.date: 12/04/2017
---

# DropDown component in Office UI Fabric

A drop-down is a list of options that is shown by clicking a drop-down button. Use a drop-down list or menu to simplify the UI design, and when users should make a choice within the UI. When the list collapses, the selected item is visible. To change the selected item, users open the list, and select a new value.
  
#### Example: Drop-down in a task pane

![An image showing the drop-down](../images/overview-with-app-dropdown.png)

## Best practices

|**Do**|**Don't**|
|:------------|:--------------|
|Use a drop-down when the default selected option is more likely to be selected than other options. By contrast, ChoiceGroup or radio buttons show all choices, thereby putting equal emphasis on all options.|Don't use a drop-down when all options are equally likely to be selected.|
|Use a drop-down when there are multiple choices that can be collapsed into one field. Also, use a drop-down for long lists of items, or when screen space is constrained.|Don’t use a drop-down if there are fewer than two choices. Instead, use a check box.|
|Use shortened statements or words in a drop-down.| |

## Variants

|**Variation**|**Description**|**Example**|
|:------------|:--------------|:----------|
|**Basic uncontrolled drop-down**|Use when many options are available for selection.|![Basic uncontrolled drop-down image](../images/dropdown-uncontrolled.png)<br/>|
|**Disabled uncontrolled drop-down with defaultSelectedKey**|Disabled state of the drop-down.|![Disabled uncontrolled drop-down with defaultSelectedKey image](../images/dropdown-disabled.png)<br/>|
|**Controlled drop-down**|Use when the default selected item is influenced by another location in your UI, and the selected item in the drop-down must be maintained.|![Controlled drop-down image](../images/dropdown-controlled.png)<br/>|

## Implementation

For details, see [Dropdown](https://dev.office.com/fabric#/components/dropdown) and [Getting started with Fabric React code sample](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact).

## See also

- [UX Design Patterns](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [Office UI Fabric in Office Add-ins](office-ui-fabric.md)
