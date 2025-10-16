---
title: Implement a pinnable task pane in an Outlook add-in
description: The task pane UX shape for add-in commands opens a vertical task pane to the right of an open message or meeting request, allowing the add-in to provide UI for more detailed interactions.
ms.date: 08/01/2025
ms.topic: how-to
ms.localizationpriority: medium
---

# Implement a pinnable task pane in Outlook

The [task pane](../design/add-in-commands.md#types-of-add-in-commands) UX shape for add-in commands opens a vertical task pane to the right of an open message or meeting request. This allows an add-in to provide UI for more detailed interactions, such as filling in multiple text fields. This task pane can be shown in the Reading Pane when viewing a list of messages, allowing for quick processing of a message.

However, by default, if a user has an add-in task pane open for a message in the Reading Pane, and then selects a new message, the task pane is automatically closed. For a heavily-used add-in, the user may prefer to keep that pane open, eliminating the need to reactivate the add-in on each message. With pinnable task panes, your add-in can give the user that option.

> [!NOTE]
> Although the pinnable task pane feature was introduced in [requirement set 1.5](/javascript/api/requirement-sets/outlook/requirement-set-1.5/outlook-requirement-set-1.5), it's currently only available to Microsoft 365 subscribers using the following:
>
> - Modern Outlook on the web
> - [New Outlook on Windows](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627)
> - Classic Outlook 2016 or later on Windows (Build 7668.2000 or later for users in the Current or Microsoft 365 Insider Channels, Build 7900.xxxx or later for users in Deferred channels)
> - Outlook on Mac (Version 16.13 (18050300) or later)

> [!IMPORTANT]
> Pinnable task panes aren't available for the following:
>
> - Appointments/Meetings
> - Outlook.com

## Supported Outlook modes

Pinnable task panes are supported in both the Message Compose and Message Read modes in Outlook. However, pinning isn't supported across different modes. This is because an add-in could have different UIs for buttons and task panes in each mode. For example, if a user pins the task pane of an add-in while reading a message, then creates a new message, they won't see the add-in's task pane from the message they're composing. To view the task pane, the user must activate the add-in from the message they're composing. If the user then pins the task pane, the task pane remains pinned the next time the user composes another message.

## Support task pane pinning

The first step is to add pinning support, which is done in the add-in manifest. The markup varies depending on the type of manifest your add-in uses.

# [Unified manifest for Microsoft 365](#tab/jsonmanifest)

[!INCLUDE [outlook-unified-manifest-mac](../includes/outlook-unified-manifest-mac.md)]

Add a `"pinnable"` property, set to `true`, to the object in the [`"actions"`](/microsoft-365/extensibility/schema/extension-runtimes-actions-item) array that defines the button or menu item that opens the task pane. The following is an example.

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

# [Add-in only manifest](#tab/xmlmanifest)

Add the [SupportsPinning](/javascript/api/manifest/action#supportspinning) element to the `<Action>` element that describes the task pane button. The following is an example.

```xml
<!-- Task pane button. -->
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

The `<SupportsPinning>` element is defined in the VersionOverrides v1.1 schema, so you will need to include a [VersionOverrides](/javascript/api/manifest/versionoverrides) element both for v1.0 and v1.1.

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
  // Update UI based on the new current item.
  updateTaskPaneUI(Office.context.mailbox.item);
}
```

> [!IMPORTANT]
> The implementation of event handlers for an `ItemChanged` event should check whether or not the Office.content.mailbox.item is null.
>
> ```js
> // Example implementation.
> function updateTaskPaneUI(item) {
>   // Assuming that item is always a read item (instead of a compose item).
>   if (item != null) console.log(item.subject);
> }
> ```

### Register the event handler

Use the [Office.context.mailbox.addHandlerAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) method to register your event handler for the `Office.EventType.ItemChanged` event. This should be done in the `Office.onReady` function of your task pane.

```js
Office.onReady(() => {
  $(document).ready(() => {
    // Set up the ItemChanged event.
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, itemChanged);
    updateTaskPaneUI(Office.context.mailbox.item);
  });
});
```

## Task pane pinning in multi-select

In Outlook on the web, on Mac, and in the new Outlook on Windows, when the task pane of an add-in that implements the [item multi-select](item-multi-select.md) feature is opened, it's automatically pinned to the Outlook client. It remains pinned even when a user switches to a different mail item or selects the **pin** icon from the task pane. The task pane can only be closed by selecting the **Close** button from the task pane.

Conversely, in classic Outlook on Windows, the task pane of a multi-select add-in isn't automatically pinned and closes when a user switches to a different mail item.

## Deploy to users

If you plan to [publish](../publish/publish.md) your Outlook add-in to [Microsoft Marketplace](https://marketplace.microsoft.com) and it's configured with a pinnable task pane, the pinned content of the add-in must not be static. That is, the pinned content must change depending on the message or appointment that's currently open or selected in the mailbox. This ensures that your add-in will pass [Microsoft Marketplace validation](/legal/marketplace/certification-policies).

## See also

For an example add-in that implements a pinnable task pane, see [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo) on GitHub.
