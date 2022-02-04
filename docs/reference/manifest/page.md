---
title: Page element in the manifest file
description: The Page element defines HTML page settings a custom function uses in Excel.
ms.date: 02/02/2022
ms.localizationpriority: medium
---

# Page element

Defines HTML page settings used by a custom function in Excel.

**Add-in type:** Custom Function

**Valid only in these VersionOverrides schemas**:

- Taskpane 1.0

For more information, see [Version overrides in the manifest](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associated with these requirement sets**:

- [CustomFunctionsRuntime 1.1](../requirement-sets/custom-functions-requirement-sets.md) 

## Attributes

None

## Child elements

|  Element  |  Required  |  Description  |
|:-----|:-----|:-----|
|  [SourceLocation](customfunctionssourcelocation.md)  |  Yes  | String with the resource id of the HTML file used by custom functions. |

## Example

```xml
<Page>
    <SourceLocation resid="pageURL"/>
</Page>
```
