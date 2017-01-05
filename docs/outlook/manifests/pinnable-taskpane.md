# Implement a pinnable taskpane in Outlook

The [taskpane](../add-in-commands-for-outlook.md#launching-a-task-pane) UX shape for add-in commands opens a vertical taskpane to the right of an open message or appointment, allowing the add-in to provide UI for more detailed interactions (filling in multiple fields, etc.). This taskpane can be shown in the Reading Pane when viewing a list of messages, allowing for quick processing of a message.

However, by default, if a user has an add-in taskpane open for a message in the Reading Pane, and then selects a new message, the task pane is automatically closed. For a heavily-used add-in, the user may prefer to keep that pane open, eliminating the need to reactivate the add-in on each message. With pinnable taskpanes, your add-in can give the user that option.

> **Note**: Pinnable taskpanes are currently only supported by Outlook 2016 (version 7628.1000).

## Enabling taskpane pinning

The first step is to enable pinning, which is done in the add-in [manifest](./manifests.md). This is done by adding the [SupportsPinning](../../../reference/manifest/action.md#supportspinning) element to the `Action` element that describes the taskpane button.

The `SupportsPinning` element is defined in the VersionOverrides v1.1 schema, so you will need to include a [VersionOverrides](../../../reference/manifest/versionoverrides.md) element both for v1.0 and v1.1.

```xml
<!-- Task pane button -->
<Control xsi:type="Button" id="msgReadOpenPaneButton">
  <Label resid="paneReadButtonLabel" />
  <Supertip>
    <Title resid="paneReadSuperTipTitle" />
    <Description resid="paneReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="green-icon-16" />
    <bt:Image size="32" resid="green-icon-32" />
    <bt:Image size="80" resid="green-icon-80" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
    <SupportsPinning>true</SupportsPinning>
  </Action>
</Control>
```

For a full example, see the `msgReadOpenPaneButton` control in the [command-demo sample manifest](https://github.com/jasonjoh/command-demo/blob/master/command-demo-manifest.xml).

## Handling UI updates based on currently selected message

To update your taskpane's UI or internal variables based on the current item, you'll need to register an event handler to get notified of the change.

### Implement the event handler

The event handler should accept a single parameter, which is an object literal. The `type` property of this object will be set to `Office.EventType.ItemChanged`. When the event is called, the `Office.context.mailbox.item` object is already updated to reflect the currently selected item.

```js
function itemChanged(eventArgs) {
  // Update UI based on the new current item
  UpdateTaskPaneUI(Office.context.mailbox.item);
}
```

### Register the event handler

Use the [Office.context.mailbox.addHandlerAsync](https://dev.outlook.com/reference/add-ins/1.5/Office.context.mailbox.html#addHandlerAsync) method to register your event handler for the `Office.EventType.ItemChanged` event. This should be done in the `Office.initialize` function for your taskpane.

```js
Office.initialize = function (reason) {
  $(document).ready(function () {

    // Set up ItemChanged event
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, itemChanged);

    UpdateTaskPaneUI(Office.context.mailbox.item);
  });
};
```

## Additional Resources

For an example add-in that implements a pinnable taskpane, see [command-demo](https://github.com/jasonjoh/command-demo) on GitHub.