---
title: Implement a pinnable task pane in an Outlook add-in
description: The task pane UX shape for add-in commands opens a vertical task pane to the right of an open message or meeting request, allowing the add-in to provide UI for more detailed interactions.
ms.date: 04/12/2024
ms.topic: how-to
ms.localizationpriority: medium
---

# Implement a pinnable task pane in Outlook

The [task pane](../design/add-in-commands.md#types-of-add-in-commands) UX shape for add-in commands opens a vertical task pane to the right of an open message or meeting request, allowing the add-in to provide UI for more detailed interactions (filling in multiple fields, etc.). This task pane can be shown in the Reading Pane when viewing a list of messages, allowing for quick processing of a message.

However, by default, if a user has an add-in task pane open for a message in the Reading Pane, and then selects a new message, the task pane is automatically closed. For a heavily-used add-in, the user may prefer to keep that pane open, eliminating the need to reactivate the add-in on each message. With pinnable task panes, your add-in can give the user that option.

> [!NOTE]
> Although the pinnable task panes feature was introduced in [requirement set 1.5](/javascript/api/requirement-sets/outlook/requirement-set-1.5/outlook-requirement-set-1.5), it's currently only available to Microsoft 365 subscribers using the following:
>
> - Outlook 2016 or later on Windows (Build 7668.2000 or later for users in the Current or Microsoft 365 Insider Channels, Build 7900.xxxx or later for users in Deferred channels)
> - Outlook 2016 or later on Mac (Version 16.13.503 or later)
> - Modern Outlook on the web

> [!IMPORTANT]
> Pinnable task panes are not available for the following:
>
> - Appointments/Meetings
> - Outlook.com

## Support task pane pinning

The first step is to add pinning support, which is done in the add-in manifest. The markup varies depending on the type of manifest.

# [Unified manifest for Microsoft 365](#tab/jsonmanifest)

Add a "pinnable" property, set to `true`, to the object in the "actions" array that defines the button or menu item that opens the task pane. The following is an example.

```json
"actions": [
    {
        "id": "OpenTaskPane",
        "type": "openPage",
        "view": "TaskPaneView",
        "displayName": "OpenTaskPane",
        "pinnable": true
    }
]
```

# [XML Manifest](#tab/xmlmanifest)

Add the [SupportsPinning](/javascript/api/manifest/action#supportspinning) element to the **\<Action\>** element that describes the task pane button. The following is an example.

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

The **\<SupportsPinning\>** element is defined in the VersionOverrides v1.1 schema, so you will need to include a [VersionOverrides](/javascript/api/manifest/versionoverrides) element both for v1.0 and v1.1.

---

For a full example, see the `msgReadOpenPaneButton` control in the [command-demo sample manifest](https://github.com/OfficeDev/outlook-add-in-command-demo/blob/master/command-demo-manifest.xml).

> [!NOTE]
> Task pane pinning is automatically supported in an add-in that activates without the Reading Pane enabled or a message first selected. To learn more, see [Activate your Outlook add-in without the Reading Pane enabled or a message selected](contextless.md).

## Handling UI updates based on currently selected message

To update your task pane's UI or internal variables based on the current item, you'll need to register an event handler to get notified of the change.

### Implement the event handler

The event handler should accept a single parameter, which is an object literal. The `type` property of this object will be set to `Office.EventType.ItemChanged`. When the event is called, the `Office.context.mailbox.item` object is already updated to reflect the currently selected item.

```js
function itemChanged(eventArgs) {
  // Update UI based on the new current item
  UpdateTaskPaneUI(Office.context.mailbox.item);
}
```

> [!IMPORTANT]
> The implementation of event handlers for an ItemChanged event should check whether or not the Office.content.mailbox.item is null.
>
> ```js
> // Example implementation
> function UpdateTaskPaneUI(item)
> {
>   // Assuming that item is always a read item (instead of a compose item).
>   if (item != null) console.log(item.subject);
> }
> ```

### Register the event handler

Use the [Office.context.mailbox.addHandlerAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) method to register your event handler for the `Office.EventType.ItemChanged` event. This should be done in the `Office.initialize` function for your task pane.

```js
Office.initialize = function (reason) {
  $(document).ready(function () {

    // Set up ItemChanged event
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, itemChanged);

    UpdateTaskPaneUI(Office.context.mailbox.item);
  });
};
```

## Deploy to users

If you plan to [publish](../publish/publish.md) your Outlook add-in to [AppSource](https://appsource.microsoft.com), and it's configured with a pinnable task pane, your add-in content must not be static and must clearly display data related to the message that is open or selected in the mailbox. This ensures that your add-in will pass [AppSource validation](/legal/marketplace/certification-policies).

## See also

For an example add-in that implements a pinnable task pane, see [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo) on GitHub.
