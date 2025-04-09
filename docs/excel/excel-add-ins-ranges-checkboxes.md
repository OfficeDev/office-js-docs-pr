---
title: Add checkboxes using the Excel JavaScript API
description: Learn how to add checkboxes using the Excel JavaScript API.
ms.date: 04/08/2025
ms.localizationpriority: medium
---

# Add checkboxes using the Excel JavaScript API

Boolean values in ranges and cells, like **TRUE** or **FALSE**, can be replaced with checkboxes using the Excel JavaScript API.

The following screenshot shows an example usage of checkboxes in a table. The table lists a variety of items, and the checkboxes indicate whether or not the items are types of fruit.

## Add checkboxes

To add checkboxes to a range, use the [`Range.control`](/javascript/api/excel/excel.range#excel-excel-range-control-member) property and set the `CellControlType` value to `checkbox`. Only boolean values, like **TRUE** or **FALSE**, will display as checkboxes in your range.

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

To check or uncheck a checkbox, change the boolean value of that checkbox.

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

To remove checkboxes from a range and return the values to simple booleans, use the [`Range.control`](/javascript/api/excel/excel.range#excel-excel-range-control-member) property and set the `CellControlType` value to `empty`.

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
