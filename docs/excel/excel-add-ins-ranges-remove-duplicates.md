---
title: Remove duplicates using the Excel JavaScript API
description: Learn how to use the Excel JavaScript API to remove duplicates.
ms.date: 02/17/2022
ms.localizationpriority: medium
---

# Remove duplicates using the Excel JavaScript API

This article provides a code sample that removes duplicate entries in a range using the Excel JavaScript API. For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).

## Remove rows with duplicate entries

The [Range.removeDuplicates](/javascript/api/excel/excel.range#excel-excel-range-removeduplicates-member(1)) method removes rows with duplicate entries in the specified columns. The method goes through each row in the range from the lowest-valued index to the highest-valued index in the range (from top to bottom). A row is deleted if a value in its specified column or columns appeared earlier in the range. Rows in the range below the deleted row are shifted up. `removeDuplicates` does not affect the position of cells outside of the range.

`removeDuplicates` takes in a `number[]` representing the column indices which are checked for duplicates. This array is zero-based and relative to the range, not the worksheet. The method also takes in a boolean parameter that specifies whether the first row is a header. When `true`, the top row is ignored when considering duplicates. The `removeDuplicates` method returns a `RemoveDuplicatesResult` object that specifies the number of rows removed and the number of unique rows remaining.

When using a range's `removeDuplicates` method, keep the following in mind.

- `removeDuplicates` considers cell values, not function results. If two different functions evaluate to the same result, the cell values are not considered duplicates.
- Empty cells are not ignored by `removeDuplicates`. The value of an empty cell is treated like any other value. This means empty rows contained within in the range will be included in the `RemoveDuplicatesResult`.

The following code sample shows the removal of entries with duplicate values in the first column.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let range = sheet.getRange("B2:D11");

    let deleteResult = range.removeDuplicates([0],true);
    deleteResult.load();

    await context.sync();

    console.log(deleteResult.removed + " entries with duplicate names removed.");
    console.log(deleteResult.uniqueRemaining + " entries with unique names remain in the range.");
});
```

### Data before duplicate entries are removed

![Data in Excel before range's remove duplicates method has been run.](../images/excel-ranges-remove-duplicates-before.png)

### Data after duplicate entries are removed

![Data in Excel after range's remove duplicates method has been run.](../images/excel-ranges-remove-duplicates-after.png)

## See also

- [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md)
- [Work with cells using the Excel JavaScript API](excel-add-ins-cells.md)
- [Cut, copy, and paste ranges using the Excel JavaScript API](excel-add-ins-ranges-cut-copy-paste.md)
- [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md)
