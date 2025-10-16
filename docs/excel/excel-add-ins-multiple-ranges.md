---
title: Work with multiple ranges simultaneously in Excel add-ins
description: Learn how the Excel JavaScript library enables your add-in to perform operations, and set properties, on multiple ranges simultaneously.
ms.date: 09/22/2025
ms.localizationpriority: medium
---

# Work with multiple ranges simultaneously in Excel add-ins

You can apply operations or set properties on several ranges at once, even if they're not contiguous. This makes code shorter and more efficient when compared to accessing each range separately.

## Key points

- Use `RangeAreas` to read or set the same thing on several separate ranges in one call.
- A property is `null` unless all member ranges share the same value.
- Set a property once on the `RangeAreas` object instead of looping, unless each range needs different logic.
- Avoid large `RangeAreas` objects made of many single cells. Narrow first with `getSpecialCells` or other filters.
- Be careful with whole columns or rows. For more details, see [Read or write to an unbounded range](excel-add-ins-ranges-unbounded.md).

## RangeAreas

A [RangeAreas](/javascript/api/excel/excel.rangeareas) object represents a set of ranges that may not touch. It shares many members with `Range`, with a few differences in how values are returned.

Examples:

- `address` returns one comma-delimited string of all addresses.
- `dataValidation` returns a single object only if every range has the same rule, otherwise it returns `null`.
- `cellCount` is the total cells across all ranges.
- `calculate` recalculates all cells in the set.
- `getEntireColumn` and `getEntireRow` return a new `RangeAreas` spanning full columns or rows for each member.
- `copyFrom` accepts either a `Range` or a `RangeAreas` as the source.

### Complete list of Range members that are also available on RangeAreas

#### Properties

Be familiar with [Read properties of RangeAreas](#read-properties-of-rangeareas) before you write code that reads any properties listed. There are subtleties to what gets returned.

- `address`
- `addressLocal`
- `cellCount`
- `conditionalFormats`
- `context`
- `dataValidation`
- `format`
- `isEntireColumn`
- `isEntireRow`
- `style`
- `worksheet`

#### Methods

- `calculate()`
- `clear()`
- `convertDataTypeToText()`
- `convertToLinkedDataType()`
- `copyFrom()`
- `getEntireColumn()`
- `getEntireRow()`
- `getIntersection()`
- `getIntersectionOrNullObject()`
- `getOffsetRange()` (named `getOffsetRangeAreas` on the `RangeAreas` object)
- `getSpecialCells()`
- `getSpecialCellsOrNullObject()`
- `getTables()`
- `getUsedRange()` (named `getUsedRangeAreas` on the `RangeAreas` object)
- `getUsedRangeOrNullObject()` (named `getUsedRangeAreasOrNullObject` on the `RangeAreas` object)
- `load()`
- `set()`
- `setDirty()`
- `toJSON()`
- `track()`
- `untrack()`

### RangeArea-specific properties and methods

The `RangeAreas` type has some properties and methods that are not on the `Range` object. The following is a selection of them.

- `areas`: A `RangeCollection` object that contains all of the ranges represented by the `RangeAreas` object. The `RangeCollection` object is also new and is similar to other Excel collection objects. It has an `items` property which is an array of `Range` objects representing the ranges.
- `areaCount`: The total number of ranges in the `RangeAreas`.
- `getOffsetRangeAreas`: Works just like [Range.getOffsetRange](/javascript/api/excel/excel.range#excel-excel-range-getoffsetrange-member(1)), except that a `RangeAreas` is returned and it contains ranges that are each offset from one of the ranges in the original `RangeAreas`.

## Create RangeAreas

You can create a `RangeAreas` object in multiple ways. The following list includes some examples.

- Call `Worksheet.getRanges()` and pass it a string with comma-delimited range addresses. If any range you want to include has been made into a [NamedItem](/javascript/api/excel/excel.nameditem), you can include the name, instead of the address, in the string.
- Call `Range.getSpecialCells()` and return a `RangeAreas` object with cells of a specific type, such as cells that contain formulas, data validation, or conditional formatting.
- Call `Workbook.getSelectedRanges()`. This method returns a `RangeAreas` representing all the ranges that are selected on the currently active worksheet.

Once you have a `RangeAreas` object, you can create others using the methods on the object that return `RangeAreas` such as `getOffsetRangeAreas` and `getIntersection`.

> [!NOTE]
> You cannot directly add additional ranges to a `RangeAreas` object. For example, the collection in `RangeAreas.areas` does not have an `add` method.

> [!WARNING]
> Do not attempt to directly add or delete members of the `RangeAreas.areas.items` array. This will lead to undesirable behavior in your code. For example, it is possible to push an additional `Range` object onto the array, but doing so will cause errors because `RangeAreas` properties and methods behave as if the new item isn't there. For example, the `areaCount` property does not include ranges pushed in this way, and the `RangeAreas.getItemAt(index)` throws an error if `index` is larger than `areasCount-1`. Similarly, deleting a `Range` object in the `RangeAreas.areas.items` array by getting a reference to it and calling its `Range.delete` method causes bugs: although the `Range` object *is* deleted, the properties and methods of the parent `RangeAreas` object behave, or try to, as if it is still in existence. For example, if your code calls `RangeAreas.calculate`, Office will try to calculate the range, but will error because the range object is gone.

## Set properties on multiple ranges

Setting a property on a `RangeAreas` object sets the corresponding property on all the ranges in the `RangeAreas.areas` collection.

The following is an example of setting a property on multiple ranges. The function highlights the ranges **F3:F5** and **H3:H5**.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let rangeAreas = sheet.getRanges("F3:F5, H3:H5");
    rangeAreas.format.fill.color = "pink";

    await context.sync();
});
```

This example applies to scenarios in which you can hard code the range addresses that you pass to `getRanges` or easily calculate them at runtime. Some of the scenarios in which this would be true include:

- The code runs in the context of a known template.
- The code runs in the context of imported data where the schema of the data is known.

## Combine `RangeAreas` with `getSpecialCells`

Filter a `RangeAreas` down to just the cells that match a criterion, such as formulas, before applying formatting or validation.

```js
await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    // Two discontiguous vertical bands.
    const targets = sheet.getRanges("A1:A100, C1:C100");

    // Narrow to only the formula cells within those bands.
    const formulaCells = targets.getSpecialCells(Excel.SpecialCellType.formulas);
    formulaCells.format.fill.color = "lightYellow";
    await context.sync();
});
```

## Get special cells from multiple ranges

The `getSpecialCells` and `getSpecialCellsOrNullObject` methods on the `RangeAreas` object work analogously to methods of the same name on the `Range` object. These methods return the cells with the specified characteristic from all of the ranges in the `RangeAreas.areas` collection. For more details on special cells, see [Find special cells within a range](excel-add-ins-ranges-special-cells.md).

When calling the `getSpecialCells` or `getSpecialCellsOrNullObject` method on a `RangeAreas` object:

- If you pass `Excel.SpecialCellType.sameConditionalFormat` as the first parameter, the method returns all cells with the same conditional formatting as the upper-leftmost cell in the first range in the `RangeAreas.areas` collection.
- If you pass `Excel.SpecialCellType.sameDataValidation` as the first parameter, the method returns all cells with the same data validation rule as the upper-leftmost cell in the first range in the `RangeAreas.areas` collection.

## Read properties of RangeAreas

Reading property values of `RangeAreas` requires care, because a given property may have different values for different ranges within the `RangeAreas`. The general rule is that if a consistent value *can* be returned it will be returned. For example, in the following code, the RGB code for pink (`#FFC0CB`) and `true` will be logged to the console because both the ranges in the `RangeAreas` object have a pink fill and both are entire columns.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();

    // The ranges are the F column and the H column.
    let rangeAreas = sheet.getRanges("F:F, H:H");  
    rangeAreas.format.fill.color = "pink";

    rangeAreas.load("format/fill/color, isEntireColumn");
    await context.sync();

    console.log(rangeAreas.format.fill.color); // #FFC0CB
    console.log(rangeAreas.isEntireColumn); // true
});
```

Since property values can differ, keep these simple rules in mind.

- Boolean properties are `true` only if they're true in all ranges, otherwise they're `false`.
- `address` always returns the comma-delimited addresses string.
- Other properties are `null` unless all ranges share the same value.

For example, the following code creates a `RangeAreas` in which only one range is an entire column and only one is filled with pink. The console will show `null` for the fill color, `false` for the `isEntireRow` property, and "Sheet1!F3:F5, Sheet1!H:H" (assuming the sheet name is "Sheet1") for the `address` property.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let rangeAreas = sheet.getRanges("F3:F5, H:H");

    let pinkColumnRange = sheet.getRange("H:H");
    pinkColumnRange.format.fill.color = "pink";

    rangeAreas.load("format/fill/color, isEntireColumn, address");
    await context.sync();

    console.log(rangeAreas.format.fill.color); // null
    console.log(rangeAreas.isEntireColumn); // false
    console.log(rangeAreas.address); // "Sheet1!F3:F5, Sheet1!H:H"
});
```

## See also

- [Fundamental programming concepts with the Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md)
- [Read or write to a large range using the Excel JavaScript API](excel-add-ins-ranges-large.md)
- [Read or write to an unbounded range using the Excel JavaScript API](excel-add-ins-ranges-unbounded.md)
