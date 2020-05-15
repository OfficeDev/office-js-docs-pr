---
title: Runtimes in the manifest file (preview)
description: The Runtimes element specifies the add-in's runtime.
ms.date: 05/15/2020
localization_priority: Normal
---

# Runtimes element (preview)

[!include[Running custom functions in browser runtime note](../../includes/excel-shared-runtime-preview-note.md)]

Specifies the runtime of your add-in. Child of the [`<Host>`](host.md) element.

In Excel, this element enables the ribbon, task pane, and custom functions to use the same runtime. For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).

In Outlook, this element enables event-based add-in activation. For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).

**Add-in type:** Task pane, Mail

> [!IMPORTANT]
> **Excel**: Shared runtime is currently in preview and only available in Excel on Windows. To try the preview features, you will need to join [Office Insider](https://insider.office.com/).
>
> **Outlook**: Event-based activation is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web. For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).

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
