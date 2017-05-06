# Button component in Office UI Fabric

Use buttons in your Office Add-in to enable users to commit changes or complete steps in a task. Make sure that the text of the button communicates the intent of the interaction. Place buttons at the bottom of the UI container of a task pane, dialog, or content pane. For example, use buttons to allow users to submit a form, close a dialog, or to move to the next page.
  
#### Example: Buttons in a task pane

![An image showing the Button](../../images/overview_withApp_button.png)

## Best Practices

|**Do**|**Don't**|
|:-----|:--------|
|Default buttons should always perform safe operations in add-ins.|Don’t place the default focus on a button that destroys data. Instead, place the focus on the button that performs the safe operation or cancels the action.|
|![Do Button example](../../images/buttonDo.png)|![Don't Button example](../../images/buttonDont.png)|
|Use only a single line of text in the label of the button. Keep text to a minimum.|Don’t put anything other than text in a button.|
|Make sure the label conveys a clear purpose of the button to the user. Use concise, specific, self-explanatory labels. Consider using a single word only.|Don’t use buttons for navigation, except for “Back” and “Next” steps. For navigation, consider using a link.|
|Expose only one or two buttons (actions) to the user. For example, “Accept” and “Cancel”. If you need to expose more actions, consider using checkboxes or radio buttons for users to select actions, and provide a single button to start the selected actions.||
|Style “Submit”, “OK”, and “Apply” buttons as primary buttons. When “Reset” or “Cancel” buttons appear alongside one of these, style them as default buttons.| |

## Variants

|**Variation**|**Description**|**Example**|
|:------------|:--------------|:----------|
|**Primary Button**|Primary buttons inherit the theme color at rest state. Use primary buttons to highlight the main call to action.|![Primary Button image](../../images/button_primary.png)|
|**Default Button**|Default buttons should always perform safe operations and should never delete.|![Default Button image](../../images/button_default.png)|
|**Compound Button**|Use compound buttons to cause actions that complete a task, or cause a transitional task.|![Compound Button image](../../images/button_compound.png)|

## Implementation

For details, see [Button](https://dev.office.com/fabric#/components/button) and [Getting started with Fabric React code sample](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact).

## Additional resources

* [UX design patterns](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
* [Office UI Fabric in Office Add-ins](office-ui-fabric.md)
