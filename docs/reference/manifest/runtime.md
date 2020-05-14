---
title: Runtime in the manifest file
description: The Runtime element configures your add-in to use a shared JavaScript runtime for its ribbon, task pane, and custom functions.
ms.date: 05/14/2020
localization_priority: Normal
---

# Runtime element

Child element of the [`<Runtimes>`](runtimes.md) element. This element configures your add-in to use a shared JavaScript runtime so that your ribbon, task pane, and custom functions, all run in the same runtime. For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).

**Add-in type:** Task pane, Mail

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
|  **lifetime="long"**  |  Yes  | Should always be `long` if you want to use a shared runtime for the Excel add-in. |
|  **resid**  |  Yes  | Specifies the URL location of the HTML page for your add-in. The `resid` must match an `id` attribute of a `Url` element in the `Resources` element. |

## See also

- [Runtimes](runtimes.md)
