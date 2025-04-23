---
title: Add checkboxes using the Excel JavaScript API
description: Learn how to add checkboxes using the Excel JavaScript API.
ms.date: 04/08/2025
ms.localizationpriority: medium
---

# Add checkboxes using the Excel JavaScript API

This article provides code samples that add, edit, and remove checkboxes from a range with the Excel JavaScript API. For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).

To use checkboxes, make sure that your range contains Boolean values, like **TRUE** or **FALSE**. Only Boolean values can be replaced with checkboxes using the Excel JavaScript API.

The following screenshot shows an example of checkboxes in a table. The table lists a variety of items, and the checkboxes indicate whether or not the items are types of fruit.

:::image type="content" source="../images/excel-range-checkbox-table.png" alt-text="A table with checkboxes in the second column.":::

> [!NOTE]
> To experiment with the code snippets in this article in a complete sample, open [Script Lab](../overview/explore-with-script-lab.md) in Excel and select [Checkboxes](https://github.com/OfficeDev/office-js-snippets/blob/prod/samples/excel/42-range/range-cell-control.yaml) in our **Samples** library.

## Add checkboxes

To add checkboxes to a range, use the [`Range.control`](/javascript/api/excel/excel.range#excel-excel-range-control-member) property to access the [`CellControl`](/javascript/api/excel/excel.cellcontrol) type, and set the `CellControlType` enum value to `checkbox`. Only Boolean values, like **TRUE** or **FALSE**, display as checkboxes in your range. The following code sample shows how to add checkboxes to the **Analysis** column of a table named **FruitTable**.

```js
await Excel.run(async (context) => {
    // This code sample shows how to add checkboxes to a table.
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    // Get the "Analysis" column in the table, without the header.
    const range = sheet.tables.getItem("FruitTable").columns.getItem("Analysis").getDataBodyRange();

    // Change the Boolean values in the range to checkboxes.
    range.control = {
      type: Excel.CellControlType.checkbox
    };
    await context.sync();
});
```

## Change the value of a checkbox

To select or clear a checkbox with the Excel JavaScript API, change the Boolean value in that cell. Use `Range.values` to change the value of a cell. The following code sample shows how to set the value of a cell to **TRUE**. Note that if the cell doesn't already display a checkbox, then the code sample simply changes the Boolean value of the cell.

```js
await Excel.run(async (context) => {
    // This code sample shows how to change the value of cell B3.
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange("B3");

    range.values = [["TRUE"]];
    await context.sync();
});
```

## Remove checkboxes

To remove checkboxes from a range and return the values to simple Booleans, use the [`Range.control`](/javascript/api/excel/excel.range#excel-excel-range-control-member) property to access the [`CellControl`](/javascript/api/excel/excel.cellcontrol) type, and set the `CellControlType` enum value to `empty`. The following code sample shows how to remove checkboxes from the **Analysis** column of a table named **FruitTable**.

```js
await Excel.run(async (context) => {
    // This code sample shows how to remove checkboxes from a table.
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    // Get the "Analysis" column in the table, without the header.
    const range = sheet.tables.getItem("FruitTable").columns.getItem("Analysis").getDataBodyRange();

    // Change the checkboxes to Boolean values.
    range.control = {
      type: Excel.CellControlType.empty
    };
    await context.sync();
});
```

> [!NOTE]
> To remove all content from a range, use the [`Range.clearOrResetContents`](/javascript/api/excel/excel.range#excel-excel-range-clearorresetcontents-member(1)) method.

## See also

- [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md)
- [Work with cells using the Excel JavaScript API](excel-add-ins-cells.md)
- [Insert a range using the Excel JavaScript API](excel-add-ins-ranges-insert.md)
