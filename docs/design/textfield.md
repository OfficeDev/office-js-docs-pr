# TextField Component in Office UI Fabric

The TextField component in add-in enables a user to type text. It's typically used to capture a single line of text but can be configured to capture multiple lines of text. The text displays on the screen in a simple, uniform format.
  
#### Example: TextField on a task pane

![An image showing the Textfield](../images/overview_withApp_textField.png)

## Best Practices

|**Do**|**Don't**|
|:------------|:--------------|
|Use the TextField to accept data input on a form or page.|Don’t use a TextField to render basic copy as part of a body element of a page.|
|Label the TextField with a helpful name.|Don’t provide an unlabeled TextField and expect that users will know what to do with it.|
|Provide concise placeholder text that specifies what content is expected to be entered.|Don’t be overly verbose with placeholder text.|
|Provide all appropriate states for the TextField (static, hover, focus, engaged, unavailable, error).|Don’t use a text-field if the valid input options can be pre-defined. Consider using a dropdown instead.|
|Provide clear designations for which fields are required vs. optional.|Don’t use a text-field for date or time entry. Consider using a datetime picker instead.|
|Whenever possible, format TextField relative to the expected entry (4-digit PIN, 10-digit phone number (3 separate fields), etc).|Don’t place a TextField inline with body copy.|

## Variants

|**Variation**|**Description**|**Example**|
|:------------|:--------------|:----------|
|**Default TextField**|Use as the default textfield.|![Default TextField image](../images/textfieldDefault.png)|
|**Disabled TextField**|Use when the textfield is inaccessable.|![Disabled TextField image](../images/textfieldDisabled.png)|
|**Required TextField**|Use when the textfield input is required.|![Required TextField image](../images/textfieldRequired.png)|
|**TextField with a placeholder**|Use when placeholder text is needed.|![TextField with a placeholder image](../images/textfieldPlaceholder.png)|
|**TextField with a placeholder**|Use when many lines of text are needed.|![TextField with a placeholder image](../images/textfieldMulti.png)|

## Implementation

For details, see [TextField](https://dev.office.com/fabric#/components/textfield) on the Office UI Fabric website.

## Additional Resources
* [UX Pattern Sample](https://office.visualstudio.com/DefaultCollection/OC/_git/GettingStarted-FabricReact)
* [GitHub Development Resources](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)