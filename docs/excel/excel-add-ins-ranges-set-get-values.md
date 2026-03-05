---
title: Set and get range values, text, or formulas using the Excel JavaScript API
description: Learn how to use the Excel JavaScript API to set and get range values, text, or formulas.
ms.date: 02/17/2022
ms.localizationpriority: medium
---

# Set and get range values, text, or formulas using the Excel JavaScript API

This article provides code samples that set and get range values, text, or formulas with the Excel JavaScript API. For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## Set values or formulas

The following code samples set values and formulas for a single cell or a range of cells.

### Set value for a single cell

The following code sample sets the value of cell **C3** to "5" and then sets the width of the columns to best fit the data.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let range = sheet.getRange("C3");
    range.values = [[ 5 ]];
    range.format.autofitColumns();

    await context.sync();
});
```

#### Data before cell value is updated

:::image type="content" source="../images/excel-ranges-set-start.png" alt-text="Data in Excel before cell value is updated.":::

#### Data after cell value is updated

:::image type="content" source="../images/excel-ranges-set-cell-value.png" alt-text="Data in Excel after cell value is updated.":::

### Set values for a range of cells

The following code sample sets values for the cells in the range **B5:D5** and then sets the width of the columns to best fit the data.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let data = [
        ["Potato Chips", 10, 1.80],
    ];

    let range = sheet.getRange("B5:D5");
    range.values = data;
    range.format.autofitColumns();

    await context.sync();
});
```

#### Data before cell values are updated

:::image type="content" source="../images/excel-ranges-set-start.png" alt-text="Data in Excel before cell values are updated.":::

#### Data after cell values are updated

:::image type="content" source="../images/excel-ranges-set-cell-values.png" alt-text="Data in Excel after cell values are updated.":::

### Set formula for a single cell

The following code sample sets a formula for cell **E3** and then sets the width of the columns to best fit the data.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let range = sheet.getRange("E3");
    range.formulas = [[ "=C3 * D3" ]];
    range.format.autofitColumns();

    await context.sync();
});
```

#### Data before cell formula is set

:::image type="content" source="../images/excel-ranges-start-set-formula.png" alt-text="Data in Excel before cell formula is set.":::

#### Data after cell formula is set

:::image type="content" source="../images/excel-ranges-set-formula.png" alt-text="Data in Excel after cell formula is set.":::

### Set formulas for a range of cells

The following code sample sets formulas for cells in the range **E2:E6** and then sets the width of the columns to best fit the data.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let data = [
        ["=C3 * D3"],
        ["=C4 * D4"],
        ["=C5 * D5"],
        ["=SUM(E3:E5)"]
    ];

    let range = sheet.getRange("E3:E6");
    range.formulas = data;
    range.format.autofitColumns();

    await context.sync();
});
```

#### Data before cell formulas are set

:::image type="content" source="../images/excel-ranges-start-set-formula.png" alt-text="Data in Excel before cell formulas are set.":::

#### Data after cell formulas are set

:::image type="content" source="../images/excel-ranges-set-formulas.png" alt-text="Data in Excel after cell formulas are set.":::

## Get values, text, or formulas

These code samples get values, text, and formulas from a range of cells.

### Get values from a range of cells

The following code sample gets the range **B2:E6**, loads its `values` property, and writes the values to the console. The `values` property of a range specifies the raw values that the cells contain. Even if some cells in a range contain formulas, the `values` property of the range specifies the raw values for those cells, not any of the formulas.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let range = sheet.getRange("B2:E6");
    range.load("values");
    await context.sync();

    console.log(JSON.stringify(range.values, null, 4));
});
```

#### Data in range (values in column E are a result of formulas)

:::image type="content" source="../images/excel-ranges-set-formulas.png" alt-text="Data in Excel after cell formulas are set.":::

#### range.values (as logged to the console by the code sample above)

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

### Get text from a range of cells

The following code sample gets the range **B2:E6**, loads its `text` property, and writes it to the console. The `text` property of a range specifies the display values for cells in the range. Even if some cells in a range contain formulas, the `text` property of the range specifies the display values for those cells, not any of the formulas.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let range = sheet.getRange("B2:E6");
    range.load("text");
    await context.sync();

    console.log(JSON.stringify(range.text, null, 4));
});
```

#### Data in range (values in column E are a result of formulas)

:::image type="content" source="../images/excel-ranges-set-formulas.png" alt-text="Data in Excel after cell formulas are set.":::

#### range.text (as logged to the console by the code sample above)

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

### Get formulas from a range of cells

The following code sample gets the range **B2:E6**, loads its `formulas` property, and writes it to the console. The `formulas` property of a range specifies the formulas for cells in the range that contain formulas and the raw values for cells in the range that do not contain formulas.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let range = sheet.getRange("B2:E6");
    range.load("formulas");
    await context.sync();

    console.log(JSON.stringify(range.formulas, null, 4));
});
```

#### Data in range (values in column E are a result of formulas)

:::image type="content" source="../images/excel-ranges-set-formulas.png" alt-text="Data in Excel after cell formulas are set.":::

#### range.formulas (as logged to the console by the code sample above)

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

## See also

- [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md)
- [Work with cells using the Excel JavaScript API](excel-add-ins-cells.md)
- [Set and get ranges using the Excel JavaScript API](excel-add-ins-ranges-set-get.md)
- [Set range format using the Excel JavaScript API](excel-add-ins-ranges-set-format.md)
