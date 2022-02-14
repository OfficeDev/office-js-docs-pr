---
title: Metadata element in the manifest file
description: The Metadata element defines the metadata settings a custom function uses in Excel.
ms.date: 02/11/2022
ms.localizationpriority: medium
---

# Metadata element

Defines the metadata settings used by a custom function in Excel.

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
|  [SourceLocation](customfunctionssourcelocation.md)  |  Yes  | String with the resource id of the JSON file used by custom functions. |

## Example

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```
