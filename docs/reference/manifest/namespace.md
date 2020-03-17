---
title: Namespace element in the manifest file
description: The Namespace element defines the namespace a custom function uses in Excel.
ms.date: 10/09/2018
localization_priority: Normal
---

# Namespace element

Defines the namespace used by a custom function in Excel.

## Attributes

|  Attribute  |  Required  |  Description  |
|:-----|:-----|:-----|
|  **resid="namespace"**  |  Yes  | Should match the ShortStrings title for your custom function, specified within the [Resources](resources.md) element. |

## Child elements

None

## Example

```xml
<Namespace resid="namespace" />
```
