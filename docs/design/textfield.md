# TextField component in Office UI Fabric

TextFields enable users to type text. It's typically used to capture a single line of text but can be configured to capture multiple lines of text. The text displays on the screen in a simple, uniform format.
  
#### Example: TextField in a task pane

![An image showing the Textfield](../../images/overview_withApp_textField.png)

## Best Practices

|**Do**|**Don't**|
|:------------|:--------------|
|Use the TextField to accept data input on a form or page.|Don’t use a TextField to render basic copy as part of a body element of a page.|
|Label the TextField with a helpful name.|Don’t use a TextField for date or time entry. Instead, use a datetime picker.|
|Use concise placeholder text to specify what content should be entered.|Don’t use TextFields if you can predefine valid input options. Instead, use a dropdown.|
|Provide all appropriate states for the TextField (static, hover, focus, engaged, unavailable, error).||
|Clearly mark required and optional fields.||
|Whenever possible, format TextFields according to the expected data format. For example, when capturing a 10-digit phone number, use 3 separate fields to store the different parts of the phone number.||

## Variants

|**Variation**|**Description**|**Example**|
|:------------|:--------------|:----------|
|**Default TextField**|Use as the default textfield.|![Default TextField image](../../images/textfieldDefault.png)|
|**Disabled TextField**|Use when the textfield is disabled.|![Disabled TextField image](../../images/textfieldDisabled.png)|
|**Required TextField**|Use when the textfield input is required.|![Required TextField image](../../images/textfieldRequired.png)|
|**TextField with a placeholder**|Use when placeholder text is needed.|![TextField with a placeholder image](../../images/textfieldPlaceholder.png)|
|**TextField with multiple lines**|Use when many lines of text are needed.|![TextField with a placeholder image](../../images/textfieldMulti.png)|

## Implementation

For details, see [TextField](https://dev.office.com/fabric#/components/textfield) and [Getting started with Fabric React code sample](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact).

## Additional resources

* [UX design patterns](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
* [Office UI Fabric in Office Add-ins](office-ui-fabric.md)
