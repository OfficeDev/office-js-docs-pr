---
title: Runtime in the manifest file
description: ''
ms.date: 01/24/2020
localization_priority: Normal
---

# Runtime element

This feature is in preview. Child element of the [`<Runtimes>`](runtimes.md) element. This element facilitates sharing of global data and function calls between Excel custom functions and the task pane of your add-in.

**Add-in type:** Task pane

## Syntax

```XML
<Runtimes>
    <Runtime resid="TaskPaneAndCustomFunction.Url" lifetime="long" />
</Runtimes>
```

## Contained in

-[Runtimes](runtimes.md)

## Attributes

|  Attribute  |  Required  |  Description  |
|:-----|:-----|:-----|
|  **lifetime="long"**  |  Yes  | Should always be listed as long if you want Excel custom functions to work while the task pane of your add-in is closed. |
|  **resid**  |  Yes  | If used for Excel custom functions, the `resid` should point to `TaskPaneAndCustomFunction.Url`. |

## See also

-[Runtimes](runtimes.md)
