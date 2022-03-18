---
title: LaunchEvent in the manifest file
description: The LaunchEvent element configures your add-in to activate based on supported events.
ms.date: 03/18/2022
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
|  **SendMode** (preview) |  No  | Used by `OnMessageSend` and `OnAppointmentSend` events. Specifies the options available to the user if your add-in stops an item from being sent or if the add-in is unavailable. If the **SendMode** property isn't included, the `SoftBlock` option is set by default. For available options, refer to [Available SendMode options](#available-sendmode-options-preview). |

## Available SendMode options (preview)

When you include the `OnMessageSend` or `OnAppointmentSend` event in the manifest, you should also set the **SendMode** property. If the **SendMode** property isn't included, the `SoftBlock` option is set by default. The following are the available options. Based on the conditions your add-in is looking for, the user is alerted if your add-in finds an issue in the item being sent.

| SendMode option | Description |
|---|---|
|`PromptUser`|If the item doesn't meet the add-in's conditions, the user can choose **Send Anyway** in the alert, or address the issue then try to send the item again. If the add-in is taking a long time to process the item, the user will be prompted with the option to stop running the add-in and choose **Send Anyway**. In the event the add-in is unavailable (for example, there's an error loading the add-in), the item will be sent.|
|`SoftBlock`|Default option if the **SendMode** property isn't included. The user is alerted that the item they're sending doesn't meet the add-in's conditions and they must address the issue before trying to send the item again. However, if the add-in is unavailable (for example, there's an error loading the add-in), the item will be sent.|
|`Block`|The item isn't sent if any of the following situations occur.<br>- The item doesn't meet the add-in's conditions.<br>- The add-in is unable to connect to the server.<br>- There's an error loading the add-in.|

## See also

- [LaunchEvents](launchevents.md)
- [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md#supported-events)
- [Use Smart Alerts and the OnMessageSend event in your Outlook add-in](../../outlook/smart-alerts-onmessagesend-walkthrough.md)
