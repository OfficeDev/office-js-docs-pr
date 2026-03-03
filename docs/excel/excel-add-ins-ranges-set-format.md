---
title: Set the format of a range using the Excel JavaScript API
description: Learn how to use the Excel JavaScript API to set the format of a range.
ms.date: 03/03/2026
ms.localizationpriority: medium
---

# Set range format using the Excel JavaScript API

This article provides code samples that set formatting for cells in a range with the Excel JavaScript API. Formatting includes fonts, colors, number formats, borders, and alignment. For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## Key points

- Use `Range.format` to access formatting properties like font, fill, borders, and alignment.
- Set `format.fill.color` and `format.font.color` using color names or hex codes.
- Use `numberFormat` to control how numbers, dates, and currency display.
- Formatting changes don't affect cell values, only their appearance.

## Set font color and fill color

The `Range.format.font` and `Range.format.fill` properties control text and background colors. Use color names like "red" or "white", or hex color codes like "#4472C4".

The following code sample sets the font color and fill color for cells in range **B2:E2**.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let range = sheet.getRange("B2:E2");
    range.format.fill.color = "#4472C4";
    range.format.font.color = "white";

    await context.sync();
});
```

### Data in range before font color and fill color are set

:::image type="content" source="../images/excel-ranges-format-before.png" alt-text="Data in Excel before format is set.":::

### Data in range after font color and fill color are set

:::image type="content" source="../images/excel-ranges-format-font-and-fill.png" alt-text="Data in Excel after format is set.":::

## Set number format

The `numberFormat` property controls how values display in cells. Number format codes follow Excel's formatting syntax. Common formats include "0.00" for decimals, "$#,##0.00" for currency, and "m/d/yyyy" for dates.

The following code sample sets the number format for the cells in range **D3:E5** to show two decimal places.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let formats = [
        ["0.00", "0.00"],
        ["0.00", "0.00"],
        ["0.00", "0.00"]
    ];

    let range = sheet.getRange("D3:E5");
    range.numberFormat = formats;

    await context.sync();
});
```

### Data in range before number format is set

:::image type="content" source="../images/excel-ranges-format-font-and-fill.png" alt-text="Data in Excel before number format is set.":::

### Data in range after number format is set

:::image type="content" source="../images/excel-ranges-format-numbers.png" alt-text="Data in Excel after number format is set.":::

## Set font properties

Set various font properties including bold, italic, size, and font name.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let range = sheet.getRange("B2:E2");

    range.format.font.bold = true;
    range.format.font.size = 14;
    range.format.font.name = "Arial";

    await context.sync();
});
```

## Set cell alignment

The `horizontalAlignment` and `verticalAlignment` properties control how content is positioned within cells.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let range = sheet.getRange("A1:E1");

    range.format.horizontalAlignment = "Center";
    range.format.verticalAlignment = "Center";

    await context.sync();
});
```

## See also

- [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md)
- [Work with cells using the Excel JavaScript API](excel-add-ins-cells.md)
- [Get a range using the Excel JavaScript API](excel-add-ins-ranges-get.md)
- [Set and get range values, text, or formulas using the Excel JavaScript API](excel-add-ins-ranges-set-get-values.md)
- [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md)
