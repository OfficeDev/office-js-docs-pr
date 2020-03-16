---
title: Metadata element in the manifest file
description: 'The Metadata element defines the metadata settings a custom function uses in Excel.'
ms.date: 10/09/2018
localization_priority: Normal
---

# Metadata element

Defines the metadata settings used by a custom function in Excel.

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
