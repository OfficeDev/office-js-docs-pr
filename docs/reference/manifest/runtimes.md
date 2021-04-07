---
title: Runtimes in the manifest file 
description: The Runtimes element specifies your add-in's runtime.
ms.date: 04/07/2021
localization_priority: Normal
---

# Runtimes element

Specifies the runtime of your add-in. Child of the [`<Host>`](host.md) element.

> [!NOTE]
> When running in Office on Windows, your add-in uses the Internet Explorer 11 browser.

**Add-in type:** Task pane, Mail

[!include[Shared JavaScript runtime requirements](../../includes/shared-runtime-requirements-note.md)]

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
- [Configure your Office Add-in to use a shared JavaScript runtime](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md)
