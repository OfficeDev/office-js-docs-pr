---
title: AllFormFactors element in the manifest file
description: Specifies the settings for an add-in for all form factors. 
ms.date: 02/25/2022
ms.localizationpriority: medium
---

# AllFormFactors element

Specifies the settings for an add-in for all form factors. Currently, the only feature using **AllFormFactors** is custom functions. **AllFormFactors** is a required element when using custom functions.

**Add-in type:** Task pane

**Valid only in these VersionOverrides schemas**:

- Task pane 1.0

For more information, see [Version overrides in the manifest](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associated with these requirement sets**:

- [CustomFunctionsRuntime 1.1](../requirement-sets/custom-functions-requirement-sets.md)

> [!NOTE]
> This element is only supported in Excel on Windows, on Mac, and on the web. It is not supported in other Office applications or on iOS or Android.

## Child elements

|  Element |  Required  |  Description  |
|:-----|:-----|:-----|
|  [ExtensionPoint](extensionpoint.md) |  Yes |  Defines where an add-in exposes functionality. |

## AllFormFactors example

```xml
<Hosts>
    <Host xsi:type="Workbook">
        <AllFormFactors>
            <ExtensionPoint xsi:type="CustomFunctions">
                    <!-- Information on this extension point -->
            </ExtensionPoint>
        </AllFormFactors>
    </Host>
</Hosts>
```
