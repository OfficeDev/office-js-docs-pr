# Toggle component in Office UI Fabric

Toggles represent a physical switch to turn things on or off. Use Toggles to present two mutually exclusive options (for example, on and off), where choosing an option results in an immediate action.
  
#### Example: Toggle in a task pane

![An image showing the Toggle](../../images/overview_withApp_toggle.png)

## Best Practices

|**Do**|**Don't**|
|:------------|:--------------|
|Use a Toggle for binary settings when changes become effective immediately after the user changes them (see image below).|Don’t use a Toggle if it requires users to perform an extra step for changes to take effect (see image below).|
|![Do Toggle example](../../images/toggleDo.png)|![Don't Toggle example](../../images/toggleDont.png)|
|Only replace the On and Off labels if there are more specific labels to use for a setting. Use short (3-4 character) labels that represent binary opposites.| |

## Variants

|**Variation**|**Description**|**Example**|
|:------------|:--------------|:----------|
|**Enabled and checked**|Use when toggled state is active.|![Enabled and checked image](../../images/toggleEnabledOn.png)|
|**Enabled and unchecked**|Use when toggled state is inactive.|![Enabled and unchecked image](../../images/toggleEnabledOff.png)|
|**Disabled and checked**|Use when the active state cannot be changed.|![Disabled and checked image](../../images/toggleDisabledOn.png)|
|**Disabled and unchecked**|Use when the inactive state cannot be changed.|![Disabled and unchecked image](../../images/toggleDisabledOff.png)|

## Implementation

For details, see [Toggle](https://dev.office.com/fabric#/components/toggle) and [Getting started with Fabric React code sample](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact).

## Additional resources

* [UX design patterns](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
* [Office UI Fabric in Office Add-ins](office-ui-fabric.md)
