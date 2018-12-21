---
title: Page element - Office Add-ins manifest
description: ''
ms.date: 10/09/2018
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
