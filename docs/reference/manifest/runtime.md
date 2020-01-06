---
title: Runtime in the manifest file
description: ''
ms.date: 01/06/2020
localization_priority: Normal
---

# Runtime element

This feature is in preview. Child element of the [`<Runtimes>`](runtime.md) element. This element facilitates sharing of global data and function calls between Excel custom functions and the task pane of your add-in. 

## Contained in

-[Runtimes](runtimes.md)

**Add-in type:** Task pane

## Syntax

```XML
<Runtimes>
    <Runtime resid="TaskPaneAndCustomFunction.Url" lifetime="long" />
</Runtimes>
```

## Attributes

|  Attribute  |  Required  |  Description  |
|:-----|:-----|:-----|
|  **lifetime="long"**  |  Yes  | Should always be listed as long if you want Excel custom functions to work while the task pane of your add-in is closed. |
|  **resid**  |  Yes  | If used for Excel custom functions, the resid should point to "TaskPaneAndCustomFunction.Url". |

## See also

-[Runtime](runtime.md)
