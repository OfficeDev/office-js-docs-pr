---
title: Runtimes in the manifest file
description: ''
ms.date: 02/14/2020
localization_priority: Normal
---

# Runtimes element

Specifies the runtime of your add-in and allows custom functions and the task pane to share global data and make function calls into each other. Child of the `<Host>` element in your manifest file.

**Add-in type:** Task pane

> [!IMPORTANT]
> Shared runtime is currently in preview and are only available on Excel on Windows. To try the preview features, you will need to join [Office Insider](https://insider.office.com/).

## Syntax

```XML
<Runtimes>
    <Runtime resid="TaskPaneAndCustomFunction.Url" lifetime="long" />
</Runtimes>
```

## Contained in 
[Host](./host.md)

## Child elements

|  Element |  Required  |  Description  |
|:-----|:-----|:-----|
|  **Runtime**     | Yes |  The Runtime for your add-in, often used with Excel custom functions.

## See also

- [Runtime](runtime.md)
