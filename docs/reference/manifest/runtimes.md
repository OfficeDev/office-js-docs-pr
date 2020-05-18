---
title: Runtimes in the manifest file 
description: 'The Runtimes element specifies the your add-in's runtime.'
ms.date: 05/17/2020
localization_priority: Normal
---
# Runtimes element

Specifies the runtime of your add-in and enables custom functions, ribbon buttons, and the task pane to use the same JavaScript runtime. Child of the `<Host>` element in your manifest file. For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).

> [!IMPORTANT]

**Add-in type:** Task pane

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
