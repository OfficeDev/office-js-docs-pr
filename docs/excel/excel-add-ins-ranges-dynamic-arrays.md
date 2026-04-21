---
title: Handle dynamic arrays and range spilling using the Excel JavaScript API
description: Learn how to handle dynamic arrays and range spilling with the Excel JavaScript API.
ms.date: 03/03/2026
ms.localizationpriority: medium
---

# Handle dynamic arrays and spilling using the Excel JavaScript API

This article provides code samples that handle dynamic arrays and range spilling using the Excel JavaScript API. A [dynamic array](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531) is an Excel feature that allows formulas to return multiple values automatically. Understanding how to work with spilled ranges programmatically enables your add-in to interact with these dynamic results effectively.

For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).

## Key points

- Dynamic array formulas automatically spill results into neighboring cells.
- Use `getSpillingToRange` to find all cells filled by a dynamic array formula.
- Use `getSpillParent` to find the original cell containing the formula that created a spilled value.
- Both methods have `*OrNullObject` versions to avoid throwing errors when no spill exists.
- Only single-cell ranges can call `getSpillParent`. Calling it on multi-cell ranges throws an error.

## Dynamic arrays

Dynamic array formulas automatically fill values into multiple cells beyond the original formula cell. This expansion is called "spilling". Common dynamic array formulas include `FILTER`, `SORT`, `UNIQUE`, `SEQUENCE`, and simple array references like `=A1:D1`.

When a formula spills, Excel automatically populates neighboring cells with the results. Your add-in can programmatically discover which cells contain these spilled values using the [Range.getSpillingToRange](/javascript/api/excel/excel.range#excel-excel-range-getspillingtorange-member(1)) method.

To handle cases where a cell might not contain a spilled formula, use the [*OrNullObject version](../develop/application-specific-api-model.md#ornullobject-methods-and-properties) `Range.getSpillingToRangeOrNullObject`. This returns the spilled range when one exists. If no spilled range exists, it returns an object whose `isNullObject` property is set to `true`. Your code then evaluates this property to determine whether the spill range exists.

### Get the spill range from a formula

The following code sample sets a formula that returns a dynamic array, then retrieves and logs the range that contains all the spilled values.

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

## Find the source of spilled values

When working with a cell that contains a spilled value, you can trace back to the original formula cell using the [Range.getSpillParent](/javascript/api/excel/excel.range#excel-excel-range-getspillparent-member(1)) method. This is useful when you need to identify which formula is responsible for populating a particular cell.

The `getSpillParent` method only works when the input range object is a single cell. Calling `getSpillParent` on a range with multiple cells throws an error. Use `Range.getSpillParentOrNullObject` to avoid an error. For more information about `*OrNullObject` methods, see [*OrNullObject methods and properties](../develop/application-specific-api-model.md#ornullobject-methods-and-properties).

### Get the formula cell from a spilled value

The following code sample shows how to find the parent cell of a spilled value. When you choose a cell that contains a spilled value and run this code, it returns the address of the spill parent.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    // Get a cell that contains a spilled value.
    let spilledCell = sheet.getRange("H4");

    // Get the parent cell whose formula is spilling into `spilledCell`.
    let spillParentRange = spilledCell.getSpillParent();
    spillParentRange.load("address");

    await context.sync();

    // Log the address of the cell containing the formula.
    console.log(`The spill parent of H4 is ${spillParentRange.address}.`);
});
```

## See also

- [Dynamic array formulas in Excel](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531)
- [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md)
- [Work with cells using the Excel JavaScript API](excel-add-ins-cells.md)
- [Set and get range values, text, or formulas using the Excel JavaScript API](excel-add-ins-ranges-set-get-values.md)
- [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md)
