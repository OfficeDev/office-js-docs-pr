---
title: Checkbox component in Office UI Fabric
description: ''
ms.date: 12/04/2017
---

# Checkbox component in Office UI Fabric

A check box is a UI element that allows users to select or clear options in add-ins. Use check boxes to allow users to select among options. Additionally, a check box may be paired with a related control. When the check box is selected or cleared, the behavior of the related control changes. For example, the related control may toggle between the visible or hidden states.
  
#### Example: Check box in a task pane

![An image showing a check box](../images/overview-with-app-checkbox.png)

## Best practices

|**Do**|**Don't**|
|:------------|:--------------|
|Use check boxes to indicate status.<br/><br/>![Do check box example](../images/checkbox-do.png)<br/>|Don’t use check boxes to show/indicate an action.<br/><br/>![Don't check box example](../images/checkbox-dont.png)<br/>|
|Use multiple check boxes when users can select multiple options, and the options are not mutually exclusive.|Don’t use a check box when users can choose only one option. When selecting only one option is required, use radio buttons.|
|Allow users to choose any combination of options when several check boxes are grouped together.|Don't put two groups of check boxes next to each other. Separate the two groups with labels.|
|Use a single check box for a secondary setting. For example, the **Remember me?** check box is a secondary setting used in a sign-in scenario.|Don’t use check boxes to turn settings on or off. To change between an on or off state, use a toggle.|

## Variants

|**Variation**|**Description**|**Example**|
|:------------|:--------------|:----------|
|**Uncontrolled check box**|Use as the default check box state. |![Uncontrolled check box image](../images/checkbox-unchecked.png)|
|**Uncontrolled check box with default checked true**|Use when the check box instance maintains its own state. |![Uncontrolled check box with default checked true image](../images/checkbox-checked.png)|
|**Disabled uncontrolled check box with default checked true**|Disabled state of the check box. |![Disabled uncontrolled check box with default checked true image](../images/checkbox-disabled.png)|
|**Controlled check box**|The checked state of this check box is decided at another location in your UI. In this scenario, the correct value is passed to the check box by an **onChange** event and re-rendering of the UI. |![Controlled check box image](../images/checkbox-unchecked.png)|

## Implementation

For details, see [Checkbox](https://dev.office.com/fabric#/components/checkbox) and [Getting started with Fabric React code sample](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact).

## Additional resources

- [UX Design Patterns](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [Office UI Fabric in Office Add-ins](office-ui-fabric.md)
