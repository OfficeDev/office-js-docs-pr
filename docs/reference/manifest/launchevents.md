---
title: LaunchEvents in the manifest file (preview)
description: The LaunchEvents element configures your add-in to activate based on supported events.
ms.date: 02/01/2021
localization_priority: Normal
---

# LaunchEvents element (preview)

Configures your add-in to activate based on supported events. Child of the [`<ExtensionPoint>`](extensionpoint.md) element. For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).

**Add-in type:** Mail

> [!IMPORTANT]
> Event-based activation is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web and on Windows. For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).

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

[ExtensionPoint](extensionpoint.md) (**LaunchEvent** mail add-in)

## Child elements

|  Element |  Required  |  Description  |
|:-----|:-----|:-----|
| [LaunchEvent](launchevent.md) | Yes |  Map supported event to its function in the JavaScript file for add-in activation. |

## See also

- [LaunchEvent](launchevent.md)
