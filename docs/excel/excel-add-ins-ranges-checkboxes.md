---
title: Add checkboxes using the Excel JavaScript API
description: Learn how to add checkboxes using the Excel JavaScript API.
ms.date: 04/08/2025
ms.localizationpriority: medium
---

# Add checkboxes using the Excel JavaScript API

## Add checkboxes

```js
await Excel.run(async (context) => {
    // Add checkboxes to the table.
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    // Get the second column in the table, without the header.
    const range = sheet.tables.getItem("FruitTable").columns.getItem("Analysis").getDataBodyRange();

    // Change the boolean values to checkboxes.
    range.control = {
      type: Excel.CellControlType.checkbox
    };
    await context.sync();
});
```

## Change the value of a checkbox

```js
await Excel.run(async (context) => {
    // Change the value of the checkbox in B3.
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange("B3");

    range.values = [["TRUE"]];
    await context.sync();
});
```

## Remove checkboxes

```js
await Excel.run(async (context) => {
    // Remove checkboxes from the table.
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    // Get the second column in the table, without the header.
    const range = sheet.tables.getItem("FruitTable").columns.getItem("Analysis").getDataBodyRange();

    // Change the checkboxes back to boolean values.
    range.control = {
      type: Excel.CellControlType.empty
    };
    await context.sync();
});
```

> [!NOTE]
> To remove all content from a range, use the `Range.clearOrResetContents` method.

## See also

- [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md)
- [Work with cells using the Excel JavaScript API](excel-add-ins-cells.md)
- [Insert a range using the Excel JavaScript API](excel-add-ins-ranges-insert.md)
