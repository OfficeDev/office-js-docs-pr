---
title: Get Excel worksheet ranges using the JavaScript API
description: Learn how to get Excel worksheet ranges by address, named range, user selection, used range, or the entire worksheet with the Excel JavaScript API.
ms.date: 06/04/2026
ms.topic: article
ms.localizationpriority: medium
ai-usage: ai-assisted
---

# Get Excel worksheet ranges with the JavaScript API

When your add-in needs to read, write, or format cells, start by getting a `Range` object. This article shows common ways to get a range in a worksheet. For the full API surface, see [Excel.Range class](/javascript/api/excel/excel.range).

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

Use the range retrieval approach that matches how your add-in identifies data.

- Use an address such as **B2:C5** when you know the exact cells.
- Use a named range when the workbook already defines a reusable name such as `MyRange`.
- Use the selected range when your add-in should operate on user-selected cells.
- Use the used range when you need the smallest area that contains data or formatting.
- Use the entire worksheet range when your add-in needs to work with every cell in the sheet.

## Get a range by address

Use `getRange(address)` when you already know the cell reference. In this example, the add-in gets **B2:C5** from the **Sample** worksheet, loads the `address` property, and writes the result to the console.

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Sample");
    const range = sheet.getRange("B2:C5");

    range.load("address");
    await context.sync();

    console.log(`The address of the range B2:C5 is "${range.address}"`);
});
```

## Get a named range

Use a named range when the worksheet already defines a meaningful name for a block of cells. In this example, the add-in gets the range named `MyRange` from the **Sample** worksheet and then reads its address.

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Sample");
    const range = sheet.getRange("MyRange");

    range.load("address");
    await context.sync();

    console.log(`The address of the range "MyRange" is "${range.address}"`);
});
```

## Get the selected range

Use `getSelectedRange()` when your add-in should work with whichever cells the user currently selects. This method is useful for actions like formatting, copying, or analyzing a user-chosen area. In this example, the add-in gets the selected range, loads its `address` property, and writes the result to the console.

```js
await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();

    range.load("address");
    await context.sync();

    console.log(`The address of the selected range is "${range.address}"`);
});
```

For more selection tasks, such as programmatically moving the selection or extending it to the edge of the used range, see [Select or get the current Excel range](excel-add-ins-ranges-set-get.md).

## Get the used range

Use `getUsedRange()` when you need the smallest range that contains any cell with a value or formatting. If the worksheet is blank, `getUsedRange()` returns a range that contains only the top-left cell. In this example, the add-in gets the used range from the **Sample** worksheet and reads its address.

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Sample");
    const range = sheet.getUsedRange();

    range.load("address");
    await context.sync();

    console.log(`The address of the used range in the worksheet is "${range.address}"`);
});
```

## Get the entire worksheet range

Use `getRange()` with no arguments when you need a range that represents the whole worksheet. In this example, the add-in gets the entire range from the **Sample** worksheet and reads its address.

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Sample");
    const range = sheet.getRange();

    range.load("address");
    await context.sync();

    console.log(`The address of the entire worksheet range is "${range.address}"`);
});
```

## Related articles

- [Core Excel object model concepts](excel-add-ins-core-concepts.md)
- [Work with cells using the Excel JavaScript API](excel-add-ins-cells.md)
- [Set and get the selected range using the Excel JavaScript API](excel-add-ins-ranges-set-get.md)
- [Set and get range values, text, or formulas using the Excel JavaScript API](excel-add-ins-ranges-set-get-values.md)
- [Find special cells within a range using the Excel JavaScript API](excel-add-ins-ranges-special-cells.md)
