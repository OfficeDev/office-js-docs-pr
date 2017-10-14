# TextField component in Office UI Fabric

A TextField enables users to type text. It's typically used to capture a single line of text but can be configured to capture multiple lines of text. The text displays on the screen in a simple, uniform format.
  
#### Example: TextField in a task pane

![An image showing the Textfield](../../images/overview_withApp_textField.png)

<br/>

## Best practices

|**Do**|**Don't**|
|:------------|:--------------|
|Use TextFields to accept data input on a form or page.|Don’t use TextFields to render basic copy as part of a body element of a page.|
|Label TextFields with a helpful names.|Don’t use TextFields for date or time entry. Instead, use a Datetime picker.|
|Use concise placeholder text to specify what content should be entered.|Don’t use TextFields if you can predefine valid input options. Instead, use a drop-down.|
|Provide all appropriate states for the TextField (static, hover, focus, engaged, unavailable, error).||
|Clearly mark required and optional fields.||
|Whenever possible, format TextFields according to the expected data format. For example, when capturing a 10-digit phone number, use three separate fields to store the different parts of the phone number.||

## Variants

|**Variation**|**Description**|**Example**|
|:------------|:--------------|:----------|
|**Default TextField**|Use as the default TextField.|![Default TextField image](../../images/textfieldDefault.png)<br/>|
|**Disabled TextField**|Use when the TextField is disabled.|![Disabled TextField image](../../images/textfieldDisabled.png)<br/>|
|**Required TextField**|Use when the TextField input is required.|![Required TextField image](../../images/textfieldRequired.png)<br/>|
|**TextField with a placeholder**|Use when placeholder text is needed.|![TextField with a placeholder image](../../images/textfieldPlaceholder.png)<br/>|
|**TextField with multiple lines**|Use when many lines of text are needed.|![TextField with a placeholder image](../../images/textfieldMulti.png)<br/>|

## Implementation

For details, see [TextField](https://dev.office.com/fabric#/components/textfield) and [Getting started with Fabric React code sample](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact).

## Additional resources

- [UX Design Patterns](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)

- [Office UI Fabric in Office Add-ins](office-ui-fabric.md)
