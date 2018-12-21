---
title: Script element in the manifest file
description: ''
ms.date: 10/09/2018
---

# Script element

Defines script settings used by a custom function in Excel.

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
