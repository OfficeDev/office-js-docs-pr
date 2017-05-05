# Dropdown Component in Office UI Fabric

A Dropdown is a list which reveals selected items by clicking a drop-down button. It is used to simplify the design and make a choice within the UI. The selected item is visible when the list is collapse. To change the value, users open the list, navigate through the list and select a new value.
  
#### Example: Dropdown on a task pane

![An image showing the Dropdown](../../images/overview_withApp_dropdown.png)

## Best Practices

|**Do**|**Don't**|
|:------------|:--------------|
|Use a Dropdown when the selected option is more likely than the alternatives (in contrast to radio buttons where all the choices are visible putting more emphasis on the other options).|Don't use dropdowns when additional options are equally likely.|
|![Do Dropdown example](../../images/dropdownDo.png)|![Don't Dropdown example](../../images/dropdownDont.png)|

|**Do**|**Don't**|
|:------------|:--------------|
|Use a Dropdown when there are multiple choices that can be collapsed under one title. Or if the list of items is long or when space is constrained.|Don’t use Dropdown if there are fewer than two choices. Use a checkbox instead.|
|Dropdowns should contain shortened statements or words.| |

## Variants

|**Variation**|**Description**|**Example**|
|:------------|:--------------|:----------|
|**Basic uncontrolled Dropdown**|Use when many options are available for selection.|![Basic uncontrolled Dropdown image](../../images/dropdownUncontrolled.png)|
|**Disabled uncontrolled Dropdown with defaultSelectedKey**|Disabled state of the Dropdown.|![Disabled uncontrolled Dropdown with defaultSelectedKey image](../../images/dropdownDisabled.png)|
|**Controlled Dropdown**|Use when the selected item is at a higher level and selection state must be maintained.|![Controlled Dropdown image](../../images/dropdownControlled.png)|

## Implementation

For details, see [Dropdown](https://dev.office.com/fabric#/components/dropdown) on the Office UI Fabric website.

## Additional Resources
* [UX Pattern Sample](https://office.visualstudio.com/DefaultCollection/OC/_git/GettingStarted-FabricReact)
* [GitHub Development Resources](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)