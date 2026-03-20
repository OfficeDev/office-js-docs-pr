---
title: Insert ranges using the Excel JavaScript API
description: Learn how to insert a range of cells with the Excel JavaScript API.
ms.date: 03/03/2026
ms.localizationpriority: medium
---

# Insert a range of cells using the Excel JavaScript API

This article provides code samples that insert a range of cells with the Excel JavaScript API. For the complete list of properties and methods that the `Range` object supports, see the [Excel.Range class](/javascript/api/excel/excel.range).

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## Key points

- Use `Range.insert` to add new cells and shift existing cells to make room.
- Specify the shift direction: down or right.
- Existing cells move to accommodate the new empty cells.
- Formulas with cell references automatically adjust after insertion.

## Insert a range of cells

The `Range.insert` method adds new empty cells to the worksheet and shifts existing cells to make room. You must specify which direction to shift the existing cells: down or right. This is useful when adding data to the middle of existing content without overwriting it.

The following code sample inserts a range of cells in location **B4:E4** and shifts other cells down to provide space for the new cells.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let range = sheet.getRange("B4:E4");

    range.insert(Excel.InsertShiftDirection.down);

    await context.sync();
});
```

### Data before range is inserted

:::image type="content" source="../images/excel-ranges-start.png" alt-text="Data in Excel before range is inserted.":::

### Data after range is inserted

:::image type="content" source="../images/excel-ranges-after-insert.png" alt-text="Data in Excel after range is inserted.":::

## See also

- [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md)
- [Work with cells using the Excel JavaScript API](excel-add-ins-cells.md)
- [Clear or delete ranges using the Excel JavaScript API](excel-add-ins-ranges-clear-delete.md)
- [Set and get range values, text, or formulas using the Excel JavaScript API](excel-add-ins-ranges-set-get-values.md)
- [Cut, copy, and paste ranges using the Excel JavaScript API](excel-add-ins-ranges-cut-copy-paste.md)
