---
title: Insert ranges using the Excel JavaScript API
description: Learn how to insert a range of cells with the Excel JavaScript API.
ms.date: 02/17/2022
ms.localizationpriority: medium
---

# Insert a range of cells using the Excel JavaScript API

This article provides a code sample that inserts a range of cells with the Excel JavaScript API. For the complete list of properties and methods that the `Range` object supports, see the [Excel.Range class](/javascript/api/excel/excel.range).

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## Insert a range of cells

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
- [Clear or delete a ranges using the Excel JavaScript API](excel-add-ins-ranges-clear-delete.md)
