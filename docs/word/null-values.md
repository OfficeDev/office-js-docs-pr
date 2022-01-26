---
title: Null values in Word add-ins
description: Learn how to work with null values in your Word add-in.
ms.date: 01/26/2022
ms.localizationpriority: medium
---

# Null values in Word add-ins

`null` has special implications in the Word JavaScript APIs. It's used to represent default values or no formatting.

## null property values in the response

Formatting properties such as [color](/javascript/api/word/word.font#word-word-font-color-member) will contain `null` values in the response when different values exist in the specified [range](/javascript/api/word/word.range). For example, if you retrieve a range and load its `range.font.color` property:

- If all text in the range has the same font color, `range.font.color` specifies that color.
- If multiple font colors are present within the range, `range.font.color` is `null`.
