---
title: AllFormFactors element in the manifest file
description: Specifies the settings for an add-in for all form factors. 
ms.date: 10/09/2018
localization_priority: Normal
---

# AllFormFactors element

Specifies the settings for an add-in for all form factors. Currently, the only feature using **AllFormFactors** is custom functions. **AllFormFactors** is a required element when using custom functions.

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
