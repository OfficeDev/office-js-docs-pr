---
title: Namespace element in the manifest file
description: The Namespace element defines the namespace a custom function uses in Excel.
ms.date: 02/11/2022
ms.localizationpriority: medium
---

# Namespace element

Defines the namespace used by a custom function in Excel.

**Add-in type:** Custom Function

**Valid only in these VersionOverrides schemas**:

- Taskpane 1.0

For more information, see [Version overrides in the manifest](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associated with these requirement sets**:

- [CustomFunctionsRuntime 1.1](../requirement-sets/custom-functions-requirement-sets.md)

## Attributes

|  Attribute  |  Required  |  Description  |
|:-----|:-----|:-----|
|  **resid="namespace"**  |  No  | Should match the ShortStrings title for your custom function, specified within the [Resources](resources.md) element. Can be no more than 32 characters. |

## Child elements

None

## Example

```xml
<Namespace resid="namespace" />
```
