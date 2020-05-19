---
title: Runtime in the manifest file
description: The Runtime element configures your add-in to use a shared JavaScript runtime for its various components, for example, ribbon, task pane, custom functions.
ms.date: 05/18/2020
localization_priority: Normal
---

# Runtime element (preview)

Configures your add-in to use a shared JavaScript runtime so that various components all run in the same runtime. Child of the [`<Runtimes>`](runtimes.md) element.

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

- [Runtimes](runtimes.md)

## Attributes

|  Attribute  |  Required  |  Description  |
|:-----|:-----|:-----|
|  **resid**  |  Yes  | Specifies the URL location of the HTML page for your add-in. The `resid` must match an `id` attribute of a `Url` element in the `Resources` element. |
|  **lifetime**  |  No  | The default value for `lifetime` is `short` and doesn't need to be specified. Outlook add-ins use only the `short` value. If you want to use a shared runtime in an Excel add-in, explicitly set the value to `long`. |

## See also

- [Runtimes](runtimes.md)
