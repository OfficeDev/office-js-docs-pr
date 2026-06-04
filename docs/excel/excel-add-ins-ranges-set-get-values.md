---
title: Set or get Excel range values, text, and formulas
description: Learn when to use `Range.values`, `Range.text`, and `Range.formulas` to write or read Excel worksheet data in an Office Add-in.
ms.date: 06/03/2026
ms.topic: article
ms.localizationpriority: medium
ai-usage: ai-assisted
---

# Set or get Excel range values, text, and formulas

This article explains how to use the Excel JavaScript APIs to write data to a worksheet or read back what a range contains. It shows when to use `Range.values`, `Range.text`, and `Range.formulas`, and how each property changes what your add-in reads or writes.

If you need to get a `Range` object before you work with its data, see [Get Excel worksheet ranges with the JavaScript API](excel-add-ins-ranges-get.md). If you also need to format the same cells, see [Set range format using the Excel JavaScript API](excel-add-ins-ranges-set-format.md).

For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## Choose the right property for range data

Use the property that matches the result your add-in needs.

- Use `range.values` to write raw values or read calculated results from cells.
- Use `range.text` to read the displayed text exactly as users see it in the worksheet.
- Use `range.formulas` to write formulas or read the formula strings from cells that contain them.

If your add-in needs to preserve some cells while updating others, or clear cells intentionally, see [Blank and null values in Excel add-ins](excel-add-ins-blank-null-values.md).

## Write values or formulas to a range

These examples show common ways to write worksheet data. Each sample gets a range, updates the target cells, and then calls `context.sync()` to apply the change.

### Write a value to one cell

Use `range.values` with a two-dimensional array, even when you write to a single cell. The following example writes `5` to cell **C3** and then auto-fits the columns.

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Sample");
    const range = sheet.getRange("C3");

    range.values = [[5]];
    range.format.autofitColumns();

    await context.sync();
});
```

#### Before the cell value is updated

:::image type="content" source="../images/excel-ranges-set-start.png" alt-text="Data in Excel before the cell value is updated.":::

#### After the cell value is updated

:::image type="content" source="../images/excel-ranges-set-cell-value.png" alt-text="Data in Excel after the cell value is updated.":::

### Write values to a range

Use a nested array to write multiple cells in one operation. In the following example, each inner array represents one row in the target range **B5:D5**.

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Sample");
    const data = [["Potato Chips", 10, 1.8]];
    const range = sheet.getRange("B5:D5");

    range.values = data;
    range.format.autofitColumns();

    await context.sync();
});
```

#### Before the cell values are updated

:::image type="content" source="../images/excel-ranges-set-start.png" alt-text="Data in Excel before the cell values are updated.":::

#### After the cell values are updated

:::image type="content" source="../images/excel-ranges-set-cell-values.png" alt-text="Data in Excel after the cell values are updated.":::

### Write a formula to one cell

Use `range.formulas` when you want Excel to calculate a result. The following example writes a formula to cell **E3** and then auto-fits the columns.

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Sample");
    const range = sheet.getRange("E3");

    range.formulas = [["=C3 * D3"]];
    range.format.autofitColumns();

    await context.sync();
});
```

#### Before the cell formula is set

:::image type="content" source="../images/excel-ranges-start-set-formula.png" alt-text="Data in Excel before the cell formula is set.":::

#### After the cell formula is set

:::image type="content" source="../images/excel-ranges-set-formula.png" alt-text="Data in Excel after the cell formula is set.":::

### Write formulas to a range

Use a two-dimensional array of formula strings to fill multiple cells at once. In the following example, the formulas in **E3:E6** calculate row totals and a grand total.

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Sample");
    const data = [
        ["=C3 * D3"],
        ["=C4 * D4"],
        ["=C5 * D5"],
        ["=SUM(E3:E5)"]
    ];
    const range = sheet.getRange("E3:E6");

    range.formulas = data;
    range.format.autofitColumns();

    await context.sync();
});
```

#### Before the cell formulas are set

:::image type="content" source="../images/excel-ranges-start-set-formula.png" alt-text="Data in Excel before the cell formulas are set.":::

#### After the cell formulas are set

:::image type="content" source="../images/excel-ranges-set-formulas.png" alt-text="Data in Excel after the cell formulas are set.":::

## Read values, displayed text, or formulas

These examples all read the same range, **B2:E6**, but each property returns different results. Review the descriptions before the code so you can choose the property that matches your scenario.

### Read raw values from a range

Use `range.values` when your add-in needs the underlying values in the cells. If a cell contains a formula, `range.values` returns the calculated result, not the formula itself.

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Sample");
    const range = sheet.getRange("B2:E6");

    range.load("values");
    await context.sync();

    console.log(JSON.stringify(range.values, null, 4));
});
```

#### Worksheet data in the range

:::image type="content" source="../images/excel-ranges-set-formulas.png" alt-text="Data in Excel after the formulas are set.":::

#### `range.values` output

```json
[
    [
        "Product",
        "Qty",
        "Unit Price",
        "Total Price"
    ],
    [
        "Almonds",
        2,
        7.5,
        15
    ],
    [
        "Coffee",
        1,
        34.5,
        34.5
    ],
    [
        "Chocolate",
        5,
        9.56,
        47.8
    ],
    [
        "",
        "",
        "",
        97.3
    ]
]
```

### Read displayed text from a range

Use `range.text` when your add-in needs the display text that users see in Excel. If a cell contains a formula, `range.text` still returns the displayed result, not the formula.

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Sample");
    const range = sheet.getRange("B2:E6");

    range.load("text");
    await context.sync();

    console.log(JSON.stringify(range.text, null, 4));
});
```

#### Worksheet data in the range

:::image type="content" source="../images/excel-ranges-set-formulas.png" alt-text="Data in Excel after the formulas are set.":::

#### `range.text` output

```json
[
    [
        "Product",
        "Qty",
        "Unit Price",
        "Total Price"
    ],
    [
        "Almonds",
        "2",
        "7.5",
        "15"
    ],
    [
        "Coffee",
        "1",
        "34.5",
        "34.5"
    ],
    [
        "Chocolate",
        "5",
        "9.56",
        "47.8"
    ],
    [
        "",
        "",
        "",
        "97.3"
    ]
]
```

### Read formulas from a range

Use `range.formulas` when your add-in needs the actual formulas from cells. For cells that don't contain formulas, `range.formulas` returns the raw value instead.

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Sample");
    const range = sheet.getRange("B2:E6");

    range.load("formulas");
    await context.sync();

    console.log(JSON.stringify(range.formulas, null, 4));
});
```

#### Worksheet data in the range

:::image type="content" source="../images/excel-ranges-set-formulas.png" alt-text="Data in Excel after the formulas are set.":::

#### `range.formulas` output

```json
[
    [
        "Product",
        "Qty",
        "Unit Price",
        "Total Price"
    ],
    [
        "Almonds",
        2,
        7.5,
        "=C3 * D3"
    ],
    [
        "Coffee",
        1,
        34.5,
        "=C4 * D4"
    ],
    [
        "Chocolate",
        5,
        9.56,
        "=C5 * D5"
    ],
    [
        "",
        "",
        "",
        "=SUM(E3:E5)"
    ]
]
```

## Related articles

- [Core Excel object model concepts](excel-add-ins-core-concepts.md)
- [Get Excel worksheet ranges with the JavaScript API](excel-add-ins-ranges-get.md)
- [Set and get the selected range using the Excel JavaScript API](excel-add-ins-ranges-set-get.md)
- [Set range format using the Excel JavaScript API](excel-add-ins-ranges-set-format.md)
- [Blank and null values in Excel add-ins](excel-add-ins-blank-null-values.md)
- [Work with cells using the Excel JavaScript API](excel-add-ins-cells.md)
