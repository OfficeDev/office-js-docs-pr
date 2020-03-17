---
title: Page element in the manifest file
description: The Page element defines HTML page settings a custom function uses in Excel.
ms.date: 10/09/2018
localization_priority: Normal
---

# Page element

Defines HTML page settings used by a custom function in Excel.

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
