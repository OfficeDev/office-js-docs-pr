---
title: Manage Excel range checkboxes with the JavaScript API
description: Add, select, clear, and remove checkboxes in Excel ranges that contain Boolean values by using the Excel JavaScript API.
ms.date: 06/03/2026
ms.topic: how-to
ms.localizationpriority: medium
ai-usage: ai-assisted
---

# Manage checkboxes in Excel ranges with the JavaScript API

Use checkboxes in Excel ranges when your add-in needs a clear yes-or-no experience, such as a task list, review tracker, or approval table. This article shows how to add checkboxes to a range, select or clear them, and remove them when you want to return to plain Boolean values.

Checkboxes only display for cells that contain Boolean values such as `true` and `false`. If you need to prepare the worksheet first, see [Set or get Excel range values, text, and formulas](excel-add-ins-ranges-set-get-values.md).

> [!NOTE]
> To try the code snippets in this article in a complete sample, open [Script Lab](../overview/explore-with-script-lab.md) in Excel and select [Checkboxes](https://github.com/OfficeDev/office-js-snippets/blob/prod/samples/excel/42-range/range-cell-control.yaml) in the **Samples** library.

## Checkbox workflow

- Store Boolean values in the target cells.
- Set [`Range.control`](/javascript/api/excel/excel.range#excel-excel-range-control-member) to `Excel.CellControlType.checkbox`.
- Write `true` or `false` to change the checkbox state.
- Set `Range.control` to `Excel.CellControlType.empty` to remove the checkboxes.

The following screenshot shows checkboxes in a table. The table lists items, and the checkboxes show whether each item is a type of fruit.

:::image type="content" source="../images/excel-range-checkbox-table.png" alt-text="A table with checkboxes in the second column.":::

## Add checkboxes to a range

In this example, the **Analysis** column of the **FruitTable** table already contains Boolean values. Setting `range.control` to `Excel.CellControlType.checkbox` turns those values into interactive checkboxes.

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    // Get the Analysis column in the table, without the header.
    const range = sheet.tables.getItem("FruitTable").columns.getItem("Analysis").getDataBodyRange();

    // Change the Boolean values in the range to checkboxes.
    range.control = {
        type: Excel.CellControlType.checkbox
    };

    await context.sync();
});
```

## Select or clear a checkbox

After a cell is configured as a checkbox, write a Boolean value to that cell to select or clear it. In the following example, writing `true` to cell **B3** selects the checkbox. Write `false` to clear it. If **B3** is not already configured as a checkbox, the cell stores the Boolean value without showing a checkbox.

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange("B3");

    range.values = [[true]];

    await context.sync();
});
```

## Remove checkboxes

Set `Range.control` to `Excel.CellControlType.empty` to remove the checkbox UI from a range and keep the underlying Boolean values. The following example removes checkboxes from the **Analysis** column of the **FruitTable** table.

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    // Get the Analysis column in the table, without the header.
    const range = sheet.tables.getItem("FruitTable").columns.getItem("Analysis").getDataBodyRange();

    // Change the checkboxes to Boolean values.
    range.control = {
        type: Excel.CellControlType.empty
    };

    await context.sync();
});
```

To remove the values along with the checkboxes, use [`Range.clearOrResetContents`](/javascript/api/excel/excel.range#excel-excel-range-clearorresetcontents-member(1)).

## See also

- [Create, read, and manage tables with the Excel JavaScript API](excel-add-ins-tables.md)
- [Set or get Excel range values, text, and formulas](excel-add-ins-ranges-set-get-values.md)
- [Work with Excel cells by using Range objects](excel-add-ins-cells.md)
- [Core Excel object model concepts for Office Add-ins](excel-add-ins-core-concepts.md)
- [Explore Office JavaScript API snippets with Script Lab](../overview/explore-with-script-lab.md)
