---
title: Work with multiple ranges simultaneously in Excel add-ins
description: Learn how the Excel JavaScript library enables your add-in to perform operations, and set properties, on multiple ranges simultaneously.
ms.date: 02/16/2022
ms.localizationpriority: medium
---

# Work with multiple ranges simultaneously in Excel add-ins

The Excel JavaScript library enables your add-in to perform operations, and set properties, on multiple ranges simultaneously. The ranges do not have to be contiguous. In addition to making your code simpler, this way of setting a property runs much faster than setting the same property individually for each of the ranges.

## RangeAreas

A set of (possibly discontiguous) ranges is represented by a [RangeAreas](/javascript/api/excel/excel.rangeareas) object. It has properties and methods similar to the `Range` type (many with the same, or similar, names), but adjustments have been made to:

- The data types for properties and the behavior of the setters and getters.
- The data types of method parameters and the method behaviors.
- The data types of method return values.

Some examples:

- `RangeAreas` has an `address` property that returns a comma-delimited string of range addresses, instead of just one address as with the `Range.address` property.
- `RangeAreas` has a `dataValidation` property that returns a `DataValidation` object that represents the data validation of all the ranges in the `RangeAreas`, if it is consistent. The property is `null` if identical `DataValidation` objects are not applied to all the all the ranges in the `RangeAreas`. This is a general, but not universal, principle with the `RangeAreas` object: *If a property does not have consistent values on all the all the ranges in the `RangeAreas`, then it is `null`.* See [Read properties of RangeAreas](#read-properties-of-rangeareas) for more information and some exceptions.
- `RangeAreas.cellCount` gets the total number of cells in all the ranges in the `RangeAreas`.
- `RangeAreas.calculate` recalculates the cells of all the ranges in the `RangeAreas`.
- `RangeAreas.getEntireColumn` and `RangeAreas.getEntireRow` return another `RangeAreas` object that represents all of the columns (or rows) in all the ranges in the `RangeAreas`. For example, if the `RangeAreas` represents "A1:C4" and "F14:L15", then `RangeAreas.getEntireColumn` returns a `RangeAreas` object that represents "A:C" and "F:L".
- `RangeAreas.copyFrom` can take either a `Range` or a `RangeAreas` parameter representing the source ranges of the copy operation.

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

You can create `RangeAreas` object in two basic ways:

- Call `Worksheet.getRanges()` and pass it a string with comma-delimited range addresses. If any range you want to include has been made into a [NamedItem](/javascript/api/excel/excel.nameditem), you can include the name, instead of the address, in the string.
- Call `Workbook.getSelectedRanges()`. This method returns a `RangeAreas` representing all the ranges that are selected on the currently active worksheet.

Once you have a `RangeAreas` object, you can create others using the methods on the object that return `RangeAreas` such as `getOffsetRangeAreas` and `getIntersection`.

> [!NOTE]
> You cannot directly add additional ranges to a `RangeAreas` object. For example, the collection in `RangeAreas.areas` does not have an `add` method.

> [!WARNING]
> Do not attempt to directly add or delete members of the the `RangeAreas.areas.items` array. This will lead to undesirable behavior in your code. For example, it is possible to push an additional `Range` object onto the array, but doing so will cause errors because `RangeAreas` properties and methods behave as if the new item isn't there. For example, the `areaCount` property does not include ranges pushed in this way, and the `RangeAreas.getItemAt(index)` throws an error if `index` is larger than `areasCount-1`. Similarly, deleting a `Range` object in the `RangeAreas.areas.items` array by getting a reference to it and calling its `Range.delete` method causes bugs: although the `Range` object *is* deleted, the properties and methods of the parent `RangeAreas` object behave, or try to, as if it is still in existence. For example, if your code calls `RangeAreas.calculate`, Office will try to calculate the range, but will error because the range object is gone.

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

Things get more complicated when consistency isn't possible. The behavior of `RangeAreas` properties follows these three principles:

- A boolean property of a `RangeAreas` object returns `false` unless the property is true for all the member ranges.
- Non-boolean properties, with the exception of the `address` property, return `null` unless the corresponding property on all the member ranges has the same value.
- The `address` property returns a comma-delimited string of the addresses of the member ranges.

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
