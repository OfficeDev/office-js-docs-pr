---
title: LaunchEvent in the manifest file
description: The LaunchEvent element configures your add-in to activate based on supported events.
ms.date: 03/16/2022
ms.localizationpriority: medium
---

# LaunchEvent element

Configures your add-in to activate based on supported events. Child of the [`<LaunchEvents>`](launchevents.md) element. For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).

**Add-in type:** Mail

**Valid only in these VersionOverrides schemas**:

- Mail 1.1

For more information, see [Version overrides in the manifest](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

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
|  **SendMode** (preview) |  No  | Used by `OnMessageSend` and `OnAppointmentSend` events. Specifies the options available to the user if your add-in stops an item from being sent or if the add-in is unable to connect to the server. If the **SendMode** property isn't included, the `SoftBlock` option is set by default. For available options, refer to [Available SendMode options](#available-sendmode-options-preview). |

## Available SendMode options (preview)

When you include the `OnMessageSend` or `OnAppointmentSend` event in the manifest, you should also set the **SendMode** property. If the **SendMode** property isn't included, the `SoftBlock` option is set by default. The following are the available options. Based on the conditions your add-in is looking for, the user is alerted if your add-in finds an issue in the item being sent.

| SendMode option | Description |
|---|---|
|`PromptUser`|In the alert, the user can choose to **Send Anyway**, or address the issue then try to send the item again.|
|`SoftBlock`|The user is alerted that the item they're sending doesn't meet the add-in conditions and the add-in will decide what to do. For example, the add-in can be set to simply alert the user about the issue or require the user to address the issue before sending the item again. If the add-in is unable to connect to the server when the add-in is activated, the item will always be sent.|
|`Block`|If the item being sent doesn't meet the add-in conditions, or if the add-in is unable to connect to the server, the item is blocked from being sent.|

## See also

- [LaunchEvents](launchevents.md)
- [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md#supported-events)
- [Use Smart Alerts and the OnMessageSend event in your Outlook add-in](../../outlook/smart-alerts-onmessagesend-walkthrough.md)
