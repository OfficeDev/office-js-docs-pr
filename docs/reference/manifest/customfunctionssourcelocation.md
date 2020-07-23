---
title: SourceLocation element in the manifest file
description: Defines the location of a resource needed by the Script or Page elements used by custom functions in Excel.
ms.date: 10/09/2018
localization_priority: Normal
---

# SourceLocation element

Defines the location of a resource needed by the Script or Page elements used by custom functions in Excel.

## Attributes

| **Attribute** | **Required** | **Description**                                                                      |
|---------------|--------------|--------------------------------------------------------------------------------------|
| resid         | Yes          | The name of a URL resource defined in the &lt;Resources&gt; section of the manifest. |

## Child elements

None

## Example

```xml
<SourceLocation resid="pageURL"/>
```
