---
title: Clear or delete ranges using the Excel JavaScript API
description: Learn how to clear or delete ranges using the Excel JavaScript API.
ms.date: 02/16/2022
ms.localizationpriority: medium
---

# Clear or delete ranges using the Excel JavaScript API

This article provides code samples that clear and delete ranges with the Excel JavaScript API. For the complete list of properties and methods supported by the `Range` object, see [Excel.Range class](/javascript/api/excel/excel.range).

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## Clear a range of cells

The following code sample clears all contents and formatting of cells in the range **E2:E5**.  

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let range = sheet.getRange("E2:E5");

    range.clear();

    await context.sync();
});
```

### Data before range is cleared

:::image type="content" source="../images/excel-ranges-start.png" alt-text="Data in Excel before range is cleared.":::

### Data after range is cleared

:::image type="content" source="../images/excel-ranges-after-clear.png" alt-text="Data in Excel after range is cleared.":::

## Delete a range of cells

The following code sample deletes the cells in the range **B4:E4** and shifts other cells up to fill the space that was vacated by the deleted cells.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let range = sheet.getRange("B4:E4");

    range.delete(Excel.DeleteShiftDirection.up);

    await context.sync();
});
```

### Data before range is deleted

:::image type="content" source="../images/excel-ranges-start.png" alt-text="Data in Excel before range is deleted.":::

### Data after range is deleted

:::image type="content" source="../images/excel-ranges-after-delete.png" alt-text="Data in Excel after range is deleted.":::

## See also

- [Work with cells using the Excel JavaScript API](excel-add-ins-cells.md)
- [Set and get ranges using the Excel JavaScript API](excel-add-ins-ranges-set-get.md)
- [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md)
