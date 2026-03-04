---
title: Cut, copy, and paste ranges using the Excel JavaScript API
description: Learn how to cut, copy, and paste ranges using the Excel JavaScript API.
ms.date: 03/03/2026
ms.localizationpriority: medium
---

# Cut, copy, and paste ranges using the Excel JavaScript API

This article provides code samples that cut, copy, and paste ranges using the Excel JavaScript API. These operations are fundamental to many Excel add-in scenarios, from duplicating templates to reorganizing data. Understanding how to use `copyFrom` and `moveTo` methods programmatically enables your add-in to replicate the familiar Excel clipboard operations users know.

For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## Key points

- Use `copyFrom` to replicate Excel's copy and paste behavior programmatically.
- Use `moveTo` to cut and paste (move) cells to a new location.
- The `copyType` parameter controls what gets copied: formulas, values, formats, or all.
- Set `skipBlanks` to `true` to preserve existing data in destination cells that correspond to blank source cells.
- Both methods work within a worksheet or across worksheets in the same workbook.

> [!TIP]
> To experiment with the cut, copy, and paste APIs from this article in a complete sample, open [Script Lab](../overview/explore-with-script-lab.md) in Excel and select [Copy and paste ranges](https://github.com/OfficeDev/office-js-snippets/blob/prod/samples/excel/42-range/range-copyfrom.yaml) in our **Samples** library.

## Copy and paste

The [Range.copyFrom](/javascript/api/excel/excel.range#excel-excel-range-copyfrom-member(1)) method replicates the **Copy** and **Paste** actions of the Excel UI. The destination is the `Range` object that `copyFrom` is called on. The source to be copied is passed as a range or a string address representing a range. You can specify only the top-left cell of the destination, and Excel automatically expands it to match the source size.

### Basic copy operation

The following code sample copies the data from **A1:E1** into the range starting at **G1** (which ends up pasting into **G1:K1**).

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    // Copy everything from "A1:E1" into "G1" and the cells afterwards ("G1:K1").
    sheet.getRange("G1").copyFrom("A1:E1");
    await context.sync();
});
```

### Copy options

`Range.copyFrom` has three optional parameters that control the copy behavior.

```TypeScript
copyFrom(sourceRange: Range | RangeAreas | string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean): void;
```

#### `copyType` parameter

The `copyType` parameter specifies what data gets copied from the source to the destination.

- `Excel.RangeCopyType.formulas` transfers the formulas in the source cells and preserves the relative positioning of those formulas' ranges. Any non-formula entries are copied as-is.
- `Excel.RangeCopyType.values` copies the data values and, in the case of formulas, the result of the formula.
- `Excel.RangeCopyType.formats` copies the formatting of the range, including font, color, and other format settings, but no values.
- `Excel.RangeCopyType.all` (the default option) copies both data and formatting, preserving cells' formulas if found.

#### `skipBlanks` parameter

The `skipBlanks` parameter sets whether blank cells are copied into the destination. When `true`, `copyFrom` skips blank cells in the source range. Skipped cells won't overwrite the existing data of their corresponding cells in the destination range. The default is `false`.

#### `transpose` parameter

The `transpose` parameter determines whether the data is transposed, meaning its rows and columns are switched, into the source location. A transposed range is flipped along the main diagonal, so rows **1**, **2**, and **3** will become columns **A**, **B**, and **C**.

### Copy with skipBlanks

The following code sample demonstrates the `skipBlanks` parameter. When `skipBlanks` is `true`, blank cells in the source don't overwrite values in the destination.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    // Copy a range, omitting the blank cells so existing data is not overwritten in those cells.
    sheet.getRange("D1").copyFrom("A1:C1",
        Excel.RangeCopyType.all,
        true, // skipBlanks
        false); // transpose
    // Copy a range, including the blank cells which will overwrite existing data in the target cells.
    sheet.getRange("D2").copyFrom("A2:C2",
        Excel.RangeCopyType.all,
        false, // skipBlanks
        false); // transpose
    await context.sync();
});
```

### Data before range is copied and pasted

:::image type="content" source="../images/excel-range-copyfrom-skipblanks-before.png" alt-text="Data in Excel before range's copy method has been run.":::

### Data after range is copied and pasted

:::image type="content" source="../images/excel-range-copyfrom-skipblanks-after.png" alt-text="Data in Excel after range's copy method has been run.":::

### Copy only values

Copying only values is useful when you want to duplicate the results of formulas without copying the formulas themselves. This is equivalent to using **Paste Special > Values** in the Excel UI.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    // Copy only the values from E3:E5 (which contain formulas) to G3.
    // The destination will contain the calculated results, not the formulas.
    sheet.getRange("G3").copyFrom("E3:E5", Excel.RangeCopyType.values);

    await context.sync();
});
```

### Copy only formatting

Copying only formatting allows you to apply a formatting template to different data. This is equivalent to using **Paste Special > Formats** in the Excel UI.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    // Copy only the formatting from A1:A5 to C1.
    // The destination will have the same font, colors, and borders, but not the values.
    sheet.getRange("C1").copyFrom("A1:A5", Excel.RangeCopyType.formats);

    await context.sync();
});
```

## Cut and paste (move) cells

The [Range.moveTo](/javascript/api/excel/excel.range#excel-excel-range-moveto-member(1)) method moves cells to a new location in the workbook. This cell movement behavior works the same as when cells are moved by [dragging the range border](https://support.microsoft.com/office/803d65eb-6a3e-4534-8c6f-ff12d1c4139e) or when taking the **Cut** and **Paste** actions. Both the formatting and values of the range are moved to the location specified as the `destinationRange` parameter.

The key difference between `moveTo` and `copyFrom` is that `moveTo` removes the content from the source location, while `copyFrom` leaves the source unchanged. Use `moveTo` when reorganizing data and `copyFrom` when duplicating it.

### Move a range

The following code sample moves a range with the `Range.moveTo` method. If the destination range is smaller than the source, it's automatically expanded to encompass the source content.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.getRange("F1").values = [["Moved Range"]];

    // Move the cells "A1:E1" to "G1" (which fills the range "G1:K1").
    sheet.getRange("A1:E1").moveTo("G1");
    await context.sync();
});
```

## See also

- [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md)
- [Work with cells using the Excel JavaScript API](excel-add-ins-cells.md)
- [Set and get range values, text, or formulas using the Excel JavaScript API](excel-add-ins-ranges-set-get-values.md)
- [Set range format using the Excel JavaScript API](excel-add-ins-ranges-set-format.md)
- [Insert a range of cells using the Excel JavaScript API](excel-add-ins-ranges-insert.md)
- [Clear or delete ranges using the Excel JavaScript API](excel-add-ins-ranges-clear-delete.md)
- [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md)
