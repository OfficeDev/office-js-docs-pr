---
title: Remove duplicates using the Excel JavaScript API
description: 'A code sample that shows how to use the Excel JavaScript API to remove duplicate values.' 
ms.date: 03/26/2021 
localization_priority: Normal
---

# Remove duplicates using the Excel JavaScript API

The [Range.removeDuplicates](/javascript/api/excel/excel.range#removeduplicates-columns--includesheader-) method removes rows with duplicate entries in the specified columns. The method goes through each row in the range from the lowest-valued index to the highest-valued index in the range (from top to bottom). A row is deleted if a value in its specified column or columns appeared earlier in the range. Rows in the range below the deleted row are shifted up. `removeDuplicates` does not affect the position of cells outside of the range.

`removeDuplicates` takes in a `number[]` representing the column indices which are checked for duplicates. This array is zero-based and relative to the range, not the worksheet. The method also takes in a boolean parameter that specifies whether the first row is a header. When **true**, the top row is ignored when considering duplicates. The `removeDuplicates` method returns a `RemoveDuplicatesResult` object that specifies the number of rows removed and the number of unique rows remaining.

When using a range's `removeDuplicates` method, keep the following in mind:

- `removeDuplicates` considers cell values, not function results. If two different functions evaluate to the same result, the cell values are not considered duplicates.
- Empty cells are not ignored by `removeDuplicates`. The value of an empty cell is treated like any other value. This means empty rows contained within in the range will be included in the `RemoveDuplicatesResult`.

The following sample shows the removal of entries with duplicate values in the first column.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:D11");

    var deleteResult = range.removeDuplicates([0],true);
    deleteResult.load();

    return context.sync().then(function () {
        console.log(deleteResult.removed + " entries with duplicate names removed.");
        console.log(deleteResult.uniqueRemaining + " entries with unique names remain in the range.");
    });
}).catch(errorHandlerFunction);
```

*Before the preceding function has been run.*

![Data in Excel before range's remove duplicates method has been run](../images/excel-ranges-remove-duplicates-before.png)

*After the preceding function has been run.*

![Data in Excel after range's remove duplicates method has been run](../images/excel-ranges-remove-duplicates-after.png)

## See also

- [Work with cells using the Excel JavaScript API](excel-add-ins-cells.md)
- [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md)
- [Cut, copy, and paste ranges using the Excel JavaScript API](excel-add-ins-cut-copy-paste.md)
- [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md)
