---
title: Runtimes in the manifest file
description: ''
ms.date: 01/06/2020
localization_priority: Normal
---

# Runtimes element

This feature is in preview. Specifies the runtime of your add-in and allows custom functions and the task pane to share global data and make function calls into each other. Should follow the `<Host>` element in your manifest file.

**Add-in type:** Task pane

## Syntax

```XML
<Runtimes>
    <Runtime resid="TaskPaneAndCustomFunction.Url" lifetime="long" />
</Runtimes>
```

## Child elements

|  Element |  Required  |  Description  |
|:-----|:-----|:-----|
|  **Runtime**     | No |  The Runtime for your add-in, often used with Excel custom functions.

## See also

-[Runtimes](runtimes.md)
