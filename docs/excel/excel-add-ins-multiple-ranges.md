---
title: Working with multiple ranges simultaneously in Excel Add-ins
description: ''
ms.date: 8/30/2018
---

# Working with multiple ranges simultaneously in Excel Add-ins (Preview)

The Excel JavaScript Library provides APIs to enable your add-in to perform operations and set properties on multiple ranges simultaneously. The ranges do not have to be contiguous with each other, or even on the same worksheet. In addition to making your code simpler, this way of setting a property runs much faster than setting the same property individually for each of the ranges.

> [!NOTE]
> The APIs described in this article require **Office 2016 Click-to-Run version 1809 Build 10820.20000** or later. (You may need to join the [Office Insider program](https://products.office.com/office-insider) to get this build.) Also, you must load the beta version of the Office JavaScript library from [Office.js CDN](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js). Finally, we don't have reference pages for these APIs yet. But the following definition type file has descriptions for them: [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).

## RangeAreas

A set of (possibly discontiguous) ranges is represented by an `Excel.RangeAreas` object. It has properties and methods that are very similar to the `Range` type (and usually have the same, or similar, names), but adjustments have been made to:

- The data types for properties, and the behavior of the setters and getters.
- The data types of method parameters and the method behaviors.
- The data types of method return values.

Some examples:

- `RangeAreas` has an `address` property that can return a comma-delimited string of range addresses, instead of just one address as on the `Range.address` property.
- `RangeAreas` has a `dataValidation` property that returns a `DataValidation` object that represents the data validation of all the ranges in the `RangeAreas`, if it is consistent. The property is `null` if identical `DataValidation` objects are not applied to all the all the ranges in the `RangeAreas`. This is a general principle with the `RangeAreas` object: *If a property does not have consistent values on all the all the ranges in the `RangeAreas`, then it is `null`.* See [Reading properties of RangeAreas](#reading-properties-of-rangeareas) for more.
- `RangeAreas.cellCount` gets the total number of cells in all the ranges in the `RangeAreas`.
- `RangeAreas.calculate` recalculates the cells of all the ranges in the `RangeAreas`.
- `RangeAreas.getEntireColumn` returns another `RangeAreas` object that represents all of the columns in all the ranges in the `RangeAreas`. For example, if the `RangeAreas` represents "A1:C4" and "F14:L15", then `RangeAreas.getEntireColumn` returns a `RangeAreas` object that represents "A:C" and "F:L".
- `RangeAreas.copyFrom` can take either a `Range` or a `RangeAreas` parameter to represent the source range(s) of the copy operation.

### Area-related properties and methods

The `RangeAreas` type has some properties and methods that are not on the `Range` object:

- `areas`: A `RangeCollection` object that contains all of the ranges represented by the `RangeAreas` object. The `RangeCollection` object is also new and is similar to other Office.js collection objects. It has an `items` property which is an array of `Range` objects representing the ranges.
- `areaCount`: The total number of ranges in the `RangeAreas`.
- `getOffsetRangeAreas`: Works just like [Range.getOffsetRange](https://docs.microsoft.com/en-us/javascript/api/excel/excel.range?view=office-js#getoffsetrange-rowoffset--columnoffset-), except that a `RangeAreas` is returned and it contains ranges that are each offset from one of the ranges in the original `RangeAreas`.
 
## Creating RangeAreas and setting properties

You can create a first `RangeAreas` object in two ways:

- Call `Worksheet.getRanges()` and pass it a string with comma-delimited range addresses. (If any range you want to include has been made into a [NamedItem](https://docs.microsoft.com/javascript/api/excel/excel.nameditem), you can include the name, instead of the address, in the string.)
- Call `Workbook.getSelectedRanges()`. This method returns a `RangeAreas` representing all the ranges that are selected when it runs.

Once you have a `RangeAreas` object, you can create others using the methods on the object that return `RangeAreas` such as `getOffsetRangeAreas` and `getIntersection`.

> [!NOTE]
> You cannot directly add additional ranges to a `RangeAreas` object. For example, the collection in `RangeAreas.areas` does not have an `add` method. It is possible to push an additional `Range` object onto the `RangeAreas.areas.items` array, but this is a bad practice because `RangeAreas` properties and methods behave as if it weren't there. For example, the `areaCount` property does include ranges pushed in this way, and the `RangeAreas.getItemAt(index)` throws an error if `index` is larger than the `areasCount-1`. Similarly, it is a bad practice to delete a `Range` object in the `RangeAreas.areas.items` array by getting a reference to it and calling its `Range.delete` method. Although the `Range` object is deleted, the properties and methods behave, or try to, as if it is still in existence.

Setting a property on a `RangeAreas` has the effect of setting the corresponding property on all the ranges in the `RangeAreas.areas` collection.  

The following is an example of setting a property on multiple ranges. Note the function is intended to highlight the ranges **F3:F5** and **H3:H5** (and no others).

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const rangeAreas = sheet.getRanges("F3:F5, H3:H5");
    rangeAreas.format.fill.color = "pink";

    return context.sync();
})
```

This example applies to scenarios in which you can hard code the range addresses that you pass to `getRanges` or easily calculate them at runtime. This would include the following scenarios, among others:

- The code runs in the context of a known template.
- The code runs in the context of imported data where the schema of the data is known.

When you can't know at coding-time which ranges you need to operate on, you must discover them at runtime. The next section discusses these scenarios. 

### Programmatically discovering range areas

The `Range.getSpecialCells()` and `Range.getSpecialCellsOrNullObject()` methods enable you to find at runtime the ranges that you want to operate on based on the characteristics of the cells and the type of the values of the cells. The following is an example. About this code, note:

- It limits the part of the sheet that needs to be searched by first calling `Worksheet.getUsedRange` and calling `getSpecialCells` for only that range.
- It passes as a parameter to `getSpecialCells` the string version of a value from the `Excel.SpecialCellType` enum. Some of the other values that could be passed instead are "Blanks" for empty cells, "Constants" for cells with literal values instead of formulas, and "SameConditionalFormat" for cells that have the same conditional formatting as the first cell in the `usedRange`. The first cell is the upper leftmost cell. For a complete list of the values in the enum, see [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).
- The `getSpecialCells` method returns a `RangeAreas` object, so all of the cells with formulas will be colored pink even if they are not all contiguous. 

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const formulaRanges = usedRange.getSpecialCells("Formulas");
    formulaRanges.format.fill.color = "pink";

    return context.sync();
})
```

In many scenarios in which you would need to discover the cells on which to operate (maybe most such scenarios), there sometimes won't be *any* cells with the targeted characteristic. If `getSpecialCells` doesn't find any, it throws an **ItemNotFound** error. This would divert the flow of control to a `catch` block/method, if there is one. If there isn't, the error just halts the function. There may be scenarios in which throwing the error is exactly what you want to happen when there are no cells with the targeted characteristic. But in scenarios in which it is normal, but perhaps uncommon, for there to be no matching cells; your code should check for this possibility and handle it gracefully without throwing an error. For these scenarios, use the `getSpecialCellsOrNullObject` method and test the `RangeAreas` object that is returned for nullity. The following is an example. Note about this code:

- The `getSpecialCellsOrNullObject` method always returns a proxy object, so it is never `null` in the ordinary JavaScript sense. But if no matching cells are found, the `isNullObject` property of the object is set to `true`.
- It calls `context.sync` *before* it tests the `isNullObject` property. This is a requirement with all `*OrNullObject` methods and properties.
- You can test this code by first selecting a range that has no formula cells and running it. Then select a range that has at least one cell with a formula and run it again.

```js
Excel.run(function (context) {
    const range = context.workbook.getSelectedRange();
    const formulaRanges = range.getSpecialCellsOrNullObject("Formulas");
    return context.sync()
        .then(function() {
            if (formulaRanges.isNullObject) {
                console.log("No cells have formulas");
            }
            else {
                formulaRanges.format.fill.color = "pink";
            }
        })
        .then(context.sync);
})
```

For simplicity, all other examples in this article will use the `getSpecialCells` method instead of  `getSpecialCellsOrNullObject`.

#### Narrow the target cells with cell value types

If your are passing either "Formulas" or "Constants" to `getSpecialCells` or `getSpecialCellsOrNullObject`, there is an optional second parameter, of enum type `Excel.SpecialCellValueType`, that can further narrow down the cells to target. You can use the parameter to specify that you only want cells having certain types of values. There are four basic types: "Error", "Logical" (which means boolean), "Numbers", and "Text". (The enum has other values besides these four which are discussed below.) The following is an example. About this code, note:

- It will highlight all, and only, cells that have a literal number value. It will not highlight cells that display a number, but have a formula; nor will highlight any boolean, text, or error state cells. 
- To test the code, be sure the worksheet has some cells with literal number values, some with other kinds of literal values, and some with formulas.

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const constantNumberRanges = usedRange.getSpecialCells("Constants", "Numbers");
    constantNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

Sometimes you may need to operate on more than one cell value type; such as all text-valued and all boolean-valued ("Logical") cells. The `Excel.SpecialCellValueType` enum has values that let you combine types. For example, "LogicalText" will target all boolean and all text-valued cells. You can combine any two or any three of the four basic types. The names of these enum values that combine basic types are always in Roman alphabetical order. So to combine error-valued, text-valued, and boolean-valued cells, use "ErrorLogicalText", not "LogicalErrorText" or "TextErrorLogical". To combine all four types, you can use "All" but this is the default, so using it is the equivalent of not using the second parameter at all. The following example, will highlight all cells with formulas that produce number or boolean values:

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const formulaLogicalNumberRanges = usedRange.getSpecialCells("Formulas", "LogicalNumbers");
    formulaLogicalNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

> [!NOTE]
> The `Excel.SpecialCellValueType` parameter can only be used if the `Excel.SpecialCellType` parameter is "Formulas" or "Constants".

### Getting RangeAreas within RangeAreas

The `RangeAreas` type itself also has `getSpecialCells` and `getSpecialCellsOrNullObject` methods which take the same two parameters. These methods return all the targeted cells from all of the ranges in the `RangeAreas.areas` collection. There is one small difference in the behavior of the methods when called on a `RangeAreas` object instead of a `Range` object: when you pass "SameConditionalFormat" as the first parameter, the method returns all cells that have the same conditional formatting as the upper leftmost cell *in the first range in the `RangeAreas.areas` collection*. The same point applies to "SameDataValidation": when passed to `Range.getSpecialCells`, it returns all cells that have the same data validation rule as the upper leftmost cell *in the range*; but when it is passed to `RangeAreas.getSpecialCells`, it returns all cells that have the same data validation rule as the upper leftmost cell *in the first range in the `RangeAreas.areas` collection*.

## Reading properties of RangeAreas

Reading property values of `RangeAreas` requires care, because a given property may have different values for different ranges within the `RangeAreas`. The general rule is that if a consistent value *can* be returned it will be returned. For example, in the following code, The RGB code for pink (`#FFC0CB`) and `true` will be logged to the console because both the ranges in the `RangeAreas` object have a pink fill and both are entire columns.

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    // The ranges are the F column and the H column.
    const rangeAreas = sheet.getRanges("F:F, H:H");  
    rangeAreas.format.fill.color = "pink";

    rangeAreas.load("format/fill/color, isEntireColumn");

    return context.sync()
        .then(function () {
            console.log(rangeAreas.format.fill.color); // #FFC0CB
            console.log(rangeAreas.isEntireColumn); // true
        })
        .then(context.sync);
})
```

Things get more complicated when consistency isn't possible. For example, the following code creates a `RangeAreas` in which only one range is an entire column and only one is filled with pink. The console will show `null` for the fill color,  `false` for the `isEntireRow` property, and "Sheet1!F3:F5, Sheet!H:H" (assuming the sheet name is "Sheet1") for the `address` property. This illustrates three principles:

- A boolean property of a `RangeAreas` object returns `false` unless the property is true for all the member ranges.
- Non-boolean properties, with the exception of the `address` property, return `null` unless the corresponding property on all the member ranges has the same value.
- The `address` property returns a comma-delimited string of the addresses of the member ranges.


```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const rangeAreas = sheet.getRanges("F3:F5, H:H");

    const pinkColumnRange = sheet.getRange("H:H");
    pinkColumnRange.format.fill.color = "pink";

    rangeAreas.load("format/fill/color, isEntireColumn, address");

    return context.sync()
        .then(function () {
            console.log(rangeAreas.format.fill.color); // null
            console.log(rangeAreas.isEntireColumn); // false
            console.log(rangeAreas.address); // "Sheet1!F3:F5, Sheet!H:H"
        })
        .then(context.sync);
})
```

## See also

- [Excel JavaScript API core concepts](https://docs.microsoft.com/javascript/office/overview/excel-add-ins-reference-overview)
- [Range Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.range)
- [RangeAreas Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.rangeareas) (This link may not work while the API is in preview.)