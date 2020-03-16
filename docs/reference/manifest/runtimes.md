---
title: Runtimes in the manifest file (preview)
description: 'The Runtimes element specifies the your add-in's runtime.'
ms.date: 02/21/2020
localization_priority: Normal
---

# Runtimes element (preview)

[!include[Running custom functions in browser runtime note](../../includes/excel-shared-runtime-preview-note.md)]

Specifies the runtime of your add-in and enables custom functions, ribbon buttons, and the task pane to use the same JavaScript runtime. Child of the `<Host>` element in your manifest file. For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).

**Add-in type:** Task pane

> [!IMPORTANT]
> Shared runtime is currently in preview and are only available on Excel on Windows. To try the preview features, you will need to join [Office Insider](https://insider.office.com/).

## Syntax

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## Contained in 
[Host](./host.md)

## Child elements

|  Element |  Required  |  Description  |
|:-----|:-----|:-----|
|  **Runtime**     | Yes |  The runtime for your add-in.

## See also

- [Runtime](runtime.md)
