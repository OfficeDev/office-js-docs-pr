---
title: Runtime in the manifest file
description: The Runtime element configures your add-in to use a shared JavaScript runtime for its various components, for example, ribbon, task pane, custom functions.
ms.date: 09/28/2021
ms.localizationpriority: medium
---

# Runtime element

Configures your add-in to use a shared JavaScript runtime so that various components all run in the same runtime. Child of the [`<Runtimes>`](runtimes.md) element.

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

- [Runtimes](runtimes.md)

## Child elements

|  Element |  Required  |  Description  |
|:-----|:-----|:-----|
| [Override](override.md) | No | **Outlook**: Specifies the URL location of the JavaScript file that Outlook Desktop requires for [LaunchEvent extension point](../../reference/manifest/extensionpoint.md#launchevent) handlers. **Important**: At present, you can only define one `<Override>` element and it must be of type `javascript`.|

## Attributes

|  Attribute  |  Required  |  Description  |
|:-----|:-----|:-----|
|  **resid**  |  Yes  | Specifies the URL location of the HTML page for your add-in. The `resid` can be no more than 32 characters and must match an `id` attribute of a `Url` element in the `Resources` element. |
|  **lifetime**  |  No  | The default value for `lifetime` is `short` and doesn't need to be specified. Outlook add-ins use only the `short` value. If you want to use a shared runtime in an Excel add-in, explicitly set the value to `long`. |

## See also

- [Runtimes](runtimes.md)
- [Configure your Office Add-in to use a shared JavaScript runtime](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md)
