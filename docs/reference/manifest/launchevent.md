---
title: LaunchEvent in the manifest file
description: The LaunchEvent element configures your add-in to activate based on supported events.
ms.date: 11/01/2021
ms.localizationpriority: medium
---

# LaunchEvent element

Configures your add-in to activate based on supported events. Child of the [`<LaunchEvents>`](launchevents.md) element. For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).

**Add-in type:** Mail

## Syntax

```XML
<ExtensionPoint xsi:type="LaunchEvent">
  <LaunchEvents>
    <LaunchEvent Type="OnNewMessageCompose" FunctionName="onMessageComposeHandler"/>
    <LaunchEvent Type="OnNewAppointmentOrganizer" FunctionName="onAppointmentComposeHandler"/>
  </LaunchEvents>
  <!-- Identifies the runtime to be used (also referenced by the Runtime element). -->
  <SourceLocation resid="WebViewRuntime.Url"/>
</ExtensionPoint>
```

## Contained in

- [LaunchEvents](launchevents.md)

## Attributes

|  Attribute  |  Required  |  Description  |
|:-----|:-----|:-----|
|  **Type**  |  Yes  | Specifies a supported event type. For the set of supported types, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md#supported-events). |
|  **FunctionName**  |  Yes  | Specifies the name of the JavaScript function to handle the event specified in the `Type` attribute. |
|  **SendMode** (preview) |  No  | Required for `OnMessageSend` and `OnAppointmentSend` events. Specifies the options available to the user if your add-in stops the item from being sent. For available options, refer to [Available SendMode options](#available-sendmode-options-preview). |

## Available SendMode options (preview)

When you include the `OnMessageSend` or `OnAppointmentSend` event in the manifest, you must also set the **SendMode** property. The following are the available options. Based on the conditions your add-in is looking for, if your add-in finds an issue in the item being sent, the user is alerted.

| SendMode option | Description |
|---|---|
|`PromptUser`|In the alert, the user can choose to **Send Anyway**, or address the issue then try to send the item again.|
|`SoftBlock`|The user must fix the issue before trying to send the item again.|

## See also

- [LaunchEvents](launchevents.md)
- [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md#supported-events)
- [Use Smart Alerts and the OnMessageSend event in your Outlook add-in](../../outlook/smart-alerts-onmessagesend-walkthrough.md)
