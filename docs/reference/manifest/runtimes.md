---
title: Runtimes in the manifest file 
description: The Runtimes element specifies your add-in's runtime.
ms.date: 09/28/2021

ms.localizationpriority: medium
---

# Runtimes element

Specifies the runtime of your add-in. Child of the [`<Host>`](host.md) element.

> [!NOTE]
> When running in Office on Windows, an add-in that has a `<Runtimes>` element in its manifest does not necessarily run in the same webview control as it otherwise would. For more information about how the versions of Windows and Office determine what webview control is normally used, see [Browsers used by Office Add-ins](../../concepts/browsers-used-by-office-web-add-ins.md). If the conditions described there for using Microsoft Edge with WebView2 (Chromium-based) are met, then the add-in uses that browser whether or not it has a `<Runtimes>` element. However, when those conditions are not met, an add-in with a `<Runtimes>` element always uses Internet Explorer 11 regardless of the Windows or Microsoft 365 version.

**Add-in type:** Task pane, Mail

**Valid only in these VersionOverrides schemas**:

 - Task pane 1.0
 - Mail 1.1

For more information, see [Version overrides in the manifest](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associated with these requirement sets**:

- [SharedRuntime 1.1](../requirement-sets/shared-runtime-requirement-sets.md) (Only when used in a task pane add-in.)

[!include[Runtimes support](../../includes/runtimes-note.md)]

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
| [Runtime](runtime.md) | Yes |  The runtime for your add-in. **Important**: At present, you can only define one `<Runtime>` element. |

## See also

- [Runtime](runtime.md)
- [Configure your Office Add-in to use a shared JavaScript runtime](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md)
