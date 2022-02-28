---
title: Script element in the manifest file
description: The Script element defines script settings a custom function uses in Excel.
ms.date: 02/11/2022
ms.localizationpriority: medium
---

# Script element

Defines script settings used by a custom function in Excel.

**Add-in type:** Custom Function

**Valid only in these VersionOverrides schemas**:

- Taskpane 1.0

For more information, see [Version overrides in the manifest](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associated with these requirement sets**:

- [CustomFunctionsRuntime 1.1](../requirement-sets/custom-functions-requirement-sets.md)

## Attributes

None

## Child elements

|Elements  |  Required  |  Description  |
|:-----|:-----|:-----|
|  [SourceLocation](customfunctionssourcelocation.md)  |  Yes  | String with resource id of the JavaScript file used by custom functions.|

## Example

```xml
<Script>
    <SourceLocation resid="scriptURL" />
    <!-- The Script element is not used in the Developer Preview. -->
</Script>
```
