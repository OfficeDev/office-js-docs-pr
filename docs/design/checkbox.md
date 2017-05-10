# Checkbox component in Office UI Fabric

A checkbox is a UI element that allows users to check or uncheck options in add-ins. Use checkboxes to allow users to select between options. Additionally, a checkbox may be paired with a related control. When the checkbox is checked or unchecked, the behavior of the related control changes. For example, the related control may toggle between the visible or hidden states.
  
#### Example: Checkbox in a task pane

![An image showing a Checkbox](../../images/overview_withApp_checkbox.png)

## Best Practices

|**Do**|**Don't**|
|:------------|:--------------|
|Use checkboxes to indicate status.|Don’t use checkboxes to show/indicate an action.|
|![Do checkbox example](../../images/checkboxDo.png)|![Don't checkbox example](../../images/checkboxDont.png)|
|Use multiple checkboxes when users can select multiple options, and the options are not mutually exclusive.|Don’t use a checkbox when users can choose only one option. Use radio buttons when selecting only one option is required.|
|Allow users to choose any combination of options when several checkboxes are grouped together.|Don't put two groups of checkboxes next to each other. Separate the two groups with labels.|
|Use a single checkbox for a secondary setting. For example, the “Remember me?” checkbox is a secondary setting used in a login scenario.|Don’t use checkboxes to turn settings on and off. To change between an on and off state, use a Toggle.|

## Variants

|**Variation**|**Description**|**Example**|
|:------------|:--------------|:----------|
|**Uncontrolled checkbox**|Use as the default checkbox state.|![Uncontrolled Checkbox image](../../images/checkbox_unchecked.png)|
|**Uncontrolled checkbox with default checked true**|Use when the checkbox instance maintains its own state|![Uncontrolled Checkbox with default checked true image](../../images/checkbox_checked.png)|
|**Disabled uncontrolled checkbox with default checked true**|Disabled state of the checkbox.|![Disabled uncontrolled Checkbox with default checked true image](../../images/checkbox_disabled.png)|
|**Controlled checkbox**|The checked state of this checkbox is decided at another location in your UI. In this scenario, the correct value is passed to the checkbox by an onChange event and re-rendering the UI.|![Controlled Checkbox image](../../images/checkbox_unchecked.png)|

## Implementation

For details, see [Checkbox](https://dev.office.com/fabric#/components/checkbox) and [Getting started with Fabric React code sample](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact).

## Additional Resources
* [UX design patterns](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
* [Office UI Fabric in Office Add-ins](office-ui-fabric.md)
