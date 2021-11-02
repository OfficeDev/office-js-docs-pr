---
title: Script element in the manifest file
description: The Script element defines script settings a custom function uses in Excel.
ms.date: 09/24/2021
ms.localizationpriority: medium
---

# Script element

Defines script settings used by a custom function in Excel.

**Add-in type:** Custom function

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
