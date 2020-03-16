---
title: Script element in the manifest file
description: The Script element defines script settings a custom function uses in Excel.
ms.date: 10/09/2018
localization_priority: Normal
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
