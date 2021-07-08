---
title: Delay execution while cell is being edited
description: 'Learn how to delay the execution of the Excel.run method when a cell is being edited.'
ms.date: 09/03/2020
localization_priority: Normal
---


# Delay execution while cell is being edited

`Excel.run` has an overload that takes in a [Excel.RunOptions](/javascript/api/excel/excel.runoptions) object. This contains a set of properties that affect platform behavior when the function runs. The following property is currently supported.

- `delayForCellEdit`: Determines whether Excel delays the batch request until the user exits cell edit mode. When **true**, the batch request is delayed and runs when the user exits cell edit mode. When **false**, the batch request automatically fails if the user is in cell edit mode (causing an error to reach the user). The default behavior with no `delayForCellEdit` property specified is equivalent to when it is **false**.

```js
Excel.run({ delayForCellEdit: true }, function (context) { ... })
```
