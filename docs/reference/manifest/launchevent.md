---
title: LaunchEvent in the manifest file (preview)
description: The LaunchEvent element configures your add-in to activate based on supported events.
ms.date: 05/15/2020
localization_priority: Normal
---

# LaunchEvent element (preview)

Configures your add-in to activate based on supported events. Child of the [`<LaunchEvents>`](launchevents.md) element. For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).

**Add-in type:** Mail

> [!IMPORTANT]
> Event-based activation is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web. For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-event-based-activation-feature).

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
|  **Type**  |  Yes  | Specifies a supported event type. Available types are `OnNewMessageCompose` and `OnNewAppointmentOrganizer`. |
|  **FunctionName**  |  Yes  | Specifies the name of the JavaScript function to handle the event specified in the `Type` attribute. |

## See also

- [LaunchEvents](launchevents.md)
