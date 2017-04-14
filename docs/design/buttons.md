# Buttons component in Office UI Fabric

Use buttons in your Office Add-in to enable a user to commit a change or complete steps in a task. Make sure that the text of the button communicates the intent of the interaction. Place buttons at the bottom of the UI container of a task pane, dialog box, or content pane.

For example, use buttons for the user to submit a form, to close a dialog box, or to move to the next settings screen to commit changes.
  
**Example: Buttons on a task pane**

![An image showing a primary and secondary button in the context of a task pane in an Office app.](../images/exampleButtonEdit@430.png)

## Best Practices

![Make sure the label conveys a clear purpose of the button to the user](../images/do1.png)
![Don’t place a button at the top of a table or inline.](../images/dont1.png)

![Use only a single line of text in the label of the button](../images/do2.png)
![Don’t place the default focus on a button that destroys data. Instead, place it on the button that performs the “safe act” and retains the content or cancels the action.](../images/dont2.png)

![Button label must descriptive of the intent action, concise, specific, self-explanatory, and usually a single word.](../images/do3.png)
![Don’t use a button to navigate to another place with exception of “Back” and “Next” buttons.](../images/dont3.png)

![Expose only one or two buttons to the user at a time. For example, “Accept” and “Cancel”.](../images/do4.png)
![Do not use buttons to toggle other UX in the same context.](../images/dont4.png)

![“Submit”, “OK”, and “Apply” buttons should always be styled as primary buttons. When “Reset” or “Cancel” buttons appear alongside one of the above, they should be styled as secondary buttons.](../images/do5.png)
![Don’t put too much text in a button – try keep text to a minimum.](../images/dont5.png)

![Default buttons should always perform safe operations in Add-in.](../images/do6.png)
![Don’t put anything other than text in a button](../images/dont6.png)

![Task buttons should be used to cause actions to complete a task or cause a transitional task.](../images/do7.png)

## Implementation

For details, see [Button](https://dev.office.com/fabric#/components/button) on the Office UI Fabric website.

## Variants

|**Variation**|**Description**|**Example**|
|:------------|:--------------|:----------|
|**Primary button**|Inherits theme color at rest state. Use this as the main call to action.| ![Primary Button Image.](../images/primary.png)|
|**Default button**|Default button should always perform safe operations and should never delete.|![Default Button Image.](../images/default.png)|
|**Compound Button**|Used to cause actions that complete a task or cause a transitional task.|![Compound Button Image.](../images/compound.png)|

For details, see [Button](https://dev.office.com/fabric#/components/button) on the Office UI Fabric website.

## Additional resources

* [UX Pattern Sample](https://office.visualstudio.com/DefaultCollection/OC/_git/GettingStarted-FabricReact)
* [GitHub Development Resources](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)