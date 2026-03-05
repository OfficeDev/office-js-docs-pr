---
title: Blank and null values in Excel add-ins
description: Learn how to work with blank and null values with the Excel JavaScript API.
ms.date: 03/03/2026
ms.localizationpriority: medium
---

# Blank and null values in Excel add-ins

`null` and empty strings have special implications in the Excel JavaScript APIs. They're used to represent empty cells, no formatting, or default values. This section details the use of `null` and empty string when getting and setting properties.

> [!TIP]
> For information about Excel JavaScript API methods that check for null objects instead of throwing `ItemNotFound` exceptions, see [*OrNullObject methods and properties](../develop/application-specific-api-model.md#ornullobject-methods-and-properties).

## Key points

- Use `null` in 2-D arrays to leave specific cells unchanged while updating others.
- `null` isn't valid for single properties like `range.values` or `range.format.fill.color`.
- Properties return `null` when a range contains multiple different values (like different font colors).
- Empty strings (`''`) clear or reset property values.
- Empty strings in responses indicate cells contain no data.

## null input in 2-D Array

In Excel, a range is represented by a 2-D array, where the first dimension is rows and the second dimension is columns. To set values, number format, or formula for only specific cells within a range, specify the values, number format, or formula for those cells in the 2-D array, and specify `null` for all other cells in the 2-D array.

For example, to update the number format for only one cell within a range, and retain the existing number format for all other cells in the range, specify the new number format for the cell to update, and specify `null` for all other cells. The following code snippet sets a new number format for the fourth cell in the range, and leaves the number format unchanged for the first three cells in the range.

```js
range.values = [['Eurasia', '29.96', '0.25', '15-Feb' ]];
range.numberFormat = [[null, null, null, 'm/d/yyyy;@']];
```

## null input for a property

`null` isn't a valid input for a single property. For example, the following code snippet is not valid, as the `values` property of the range cannot be set to `null`.

```js
range.values = null; // This is not a valid snippet. 
```

Likewise, the following code snippet isn't valid, as `null` isn't a valid value for the `color` property.

```js
range.format.fill.color =  null;  // This is not a valid snippet. 
```

## null property values in the response

Formatting properties such as `size` and `color` contain `null` values in the response when different values exist in the specified range. For example, if you retrieve a range and load its `format.font.color` property:

- If all cells in the range have the same font color, then `range.format.font.color` specifies that color.
- If multiple font colors are present within the range, then `range.format.font.color` is `null`.

## Blank input for a property

When you specify a blank value for a property (such as two quotation marks with no space in-between `''`), it's interpreted as an instruction to clear or reset the property. For example:

- If you specify a blank value for the `values` property of a range, the content of the range is cleared.
- If you specify a blank value for the `numberFormat` property, the number format is reset to `General`.
- If you specify a blank value for the `formulas` property and `formulasLocals` property, the formula values are cleared.

The following code sample shows clearing values from a range.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let range = sheet.getRange("A1:D1");

    // Clear the values in the range.
    range.values = [['', '', '', '']];

    await context.sync();
});
```

## Blank property values in the response

For read operations, a blank property value in the response (such as two quotation marks with no space in-between `''`) indicates that cell contains no data or value. In the first example below, the first and last cell in the range contain no data. In the second example, the first two cells in the range do not contain a formula.

```js
range.values = [['', 'some', 'data', 'in', 'other', 'cells', '']];
```

```js
range.formulas = [['', '', '=Rand()']];
```

## See also

- [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md)
- [Work with cells using the Excel JavaScript API](excel-add-ins-cells.md)
- [Set and get range values, text, or formulas using the Excel JavaScript API](excel-add-ins-ranges-set-get-values.md)
- [Set range format using the Excel JavaScript API](excel-add-ins-ranges-set-format.md)
