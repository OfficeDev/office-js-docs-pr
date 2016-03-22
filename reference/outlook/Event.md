 

# Event

The `event` object is passed as a parameter to add-in functions invoked by UI-less command buttons. The object allows the add-in to identify which button was clicked and to signal the host that it has completed its processing.

For example, consider a button defined in an add-in manifest as follows:

```
<Control xsi:type="Button" id="eventTestButton">
  <Label resid="eventButtonLabel" />
  <Tooltip resid="eventButtonTooltip" />
  <Supertip>
    <Title resid="eventSuperTipTitle" />
    <Description resid="eventSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="blue-icon-16" />
    <bt:Image size="32" resid="blue-icon-32" />
    <bt:Image size="80" resid="blue-icon-80" />
  </Icon>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>testEventObject</FunctionName>
  </Action>
</Control>
```

The button has an `id` attribute set to `eventTestButton`, and will invoke the `testEventObject` function defined in the add-in. That function looks like this:

```
function testEventObject(event) {
  // The event object implements the Event interface

  // This value will be "eventTestButton"
  var buttonId = event.source.id;

  // Signal to the host app that processing is complete.
  event.completed();
}
```

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](./tutorial-api-requirement-sets.md)| 1.3|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| Restricted|
|Applicable Outlook mode| Compose or read|

### Members

####  source :Object

Gets the identifier of the add-in command button that invoked the method.

The `source` property returns an object with the following properties.

| Property | Description |
| --- | --- |
| `id` | The value of the `id` attribute of the `Control` element that defines the add-in command button in the add-in manifest. |

This value can be used when more than one button invokes the same function, but you need to take different actions based on which button was clicked.

##### Type:

*   Object

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](./tutorial-api-requirement-sets.md)| 1.3|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| Restricted|
|Applicable Outlook mode| Compose or read|

##### Example

```
// Function is used by two buttons:
// button1 and button2
function multiButton (event) {
  // Check which button was clicked
  var buttonId = event.source.id;

  if (buttonId === 'button1') {
    doButton1Action();
  else {
    doButton2Action();
  }

  event.completed();
}
```

### Methods

####  completed()

Indicates that the add-in has completed processing that was triggered by an add-in command button.

This method must be called at the end of a function which was invoked by an add-in command defined with an `Action` element with an `xsi:type` attribute set to `ExecuteFunction`. Calling this method signals the host client that the function is complete and that it can clean up any state involved with invoking the function. For example, if the user closes Outlook before this method is called, Outlook will warn that a function is still executing.

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](./tutorial-api-requirement-sets.md)| 1.3|
|[Minimum permission level](https://msdn.microsoft.com/EN-US/library/office/fp161087.aspx)| Restricted|
|Applicable Outlook mode| Compose or read|

##### Example

```
function processItem (event) {
  // Do some processing

  event.completed();
}
```