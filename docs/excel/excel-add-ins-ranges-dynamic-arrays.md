---
title: Handle dynamic arrays and range spilling using the Excel JavaScript API
description: Learn how to handle dynamic arrays and range spilling with the Excel JavaScript API.
ms.date: 02/17/2022
ms.localizationpriority: medium
---

# Handle dynamic arrays and spilling using the Excel JavaScript API

This article provides a code sample that handles dynamic arrays and range spilling using the Excel JavaScript API. For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).

## Dynamic arrays

Some Excel formulas return [Dynamic arrays](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531). These fill the values of multiple cells outside of the formula's original cell. This value overflow is referred to as a "spill". Your add-in can find the range used for a spill with the [Range.getSpillingToRange](/javascript/api/excel/excel.range#excel-excel-range-getspillingtorange-member(1)) method. There is also a [*OrNullObject version](../develop/application-specific-api-model.md#ornullobject-methods-and-properties), `Range.getSpillingToRangeOrNullObject`.

The following sample shows a basic formula that copies the contents of a range into a cell, which spills into neighboring cells. The add-in then logs the range that contains the spill.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    // Set G4 to a formula that returns a dynamic array.
    let targetCell = sheet.getRange("G4");
    targetCell.formulas = [["=A4:D4"]];

    // Get the address of the cells that the dynamic array spilled into.
    let spillRange = targetCell.getSpillingToRange();
    spillRange.load("address");

    // Sync and log the spilled-to range.
    await context.sync();

    // This will log the range as "G4:J4".
    console.log(`Copying the table headers spilled into ${spillRange.address}.`);
});
```

## Range spilling

Find the cell responsible for spilling into a given cell by using the [Range.getSpillParent](/javascript/api/excel/excel.range#excel-excel-range-getspillparent-member(1)) method. Note that `getSpillParent` only works when the range object is a single cell. Calling `getSpillParent` on a range with multiple cells will result in an error being thrown (or a null range being returned for `Range.getSpillParentOrNullObject`).

## See also

- [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md)
- [Work with cells using the Excel JavaScript API](excel-add-ins-cells.md)
- [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md)
