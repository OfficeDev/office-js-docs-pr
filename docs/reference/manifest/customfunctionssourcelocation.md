---
title: SourceLocation element for custom functions in the manifest file
description: Defines the location of a resource needed by the Script or Page elements used by custom functions in Excel.
ms.date: 08/07/2020
ms.localizationpriority: medium
---

# SourceLocation element (custom functions)

Defines the location of a resource needed by the Script or Page elements used by custom functions in Excel.

## Attributes

| Attribute | Required | Description                                                                          |
|-----------|----------|--------------------------------------------------------------------------------------|
| resid     | Yes      | The name of a URL resource defined in the &lt;Resources&gt; section of the manifest. Can be no more than 32 characters. |

## Child elements

None

## Example

```xml
<SourceLocation resid="pageURL"/>
```
