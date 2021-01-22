---
title: Runtimes in the manifest file 
description: The Runtimes element specifies your add-in's runtime.
ms.date: 02/01/2021
localization_priority: Normal
---
# Runtimes element

Specifies the runtime of your add-in. Child of the [`<Host>`](host.md) element.

> [!NOTE]
> When running in Office on Windows, your add-in uses the Internet Explorer 11 browser.

In Excel, this element enables the ribbon, task pane, and custom functions to use the same runtime. For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../develop/configure-your-add-in-to-use-a-shared-runtime.md).

In Outlook, this element enables event-based add-in activation. For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).

**Add-in type:** Task pane, Mail

> [!IMPORTANT]
> **Outlook**: The event-based activation feature is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web and Windows. For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).

## Syntax

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## Contained in

[Host](host.md)

## Child elements

|  Element |  Required  |  Description  |
|:-----|:-----|:-----|
| [Runtime](runtime.md) | Yes |  The runtime for your add-in. |

## See also

- [Runtime](runtime.md)
