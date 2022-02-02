---
title: SourceLocation element for custom functions in the manifest file
description: Defines the location of a resource needed by the Script or Page elements used by custom functions in Excel.
ms.date: 02/02/2022
ms.localizationpriority: medium
---

# SourceLocation element (custom functions)

Defines the location of a resource needed by the **Script** or **Page** elements used by custom functions in Excel.

> [!IMPORTANT]
> This article refers only to the **SourceLocation** that is a child of the **Page** or **Script** elements. See [SourceLocation](sourcelocation.md) for information about the **SourceLocation** element of the base manifest.

**Add-in type:** Custom function

**Valid only in these VersionOverrides schemas**:

- Taskpane 1.0

For more information, see [Version overrides in the manifest](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associated with these requirement sets**:

- [CustomFunctionsRuntime 1.1](../requirement-sets/custom-functions-requirement-sets.md)

## Attributes

| Attribute | Required | Description                                                                          |
|-----------|----------|--------------------------------------------------------------------------------------|
| resid     | Yes      | The name of a URL resource defined in the **Resources** section of the manifest. Can be no more than 32 characters. |

## Child elements

None

## Example

```xml
<SourceLocation resid="pageURL"/>
```
