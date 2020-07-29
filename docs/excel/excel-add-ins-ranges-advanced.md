---
title: Work with ranges using the Excel JavaScript API (advanced)
description: 'Advanced range object functions and scenarios, such as special cells, remove duplicates, and working with dates.' 
ms.date: 05/06/2020 
localization_priority: Normal
---

# Work with ranges using the Excel JavaScript API (advanced)

This article builds upon information in [Work with ranges using the Excel JavaScript API (fundamental)](excel-add-ins-ranges.md) by providing code samples that show how to perform more advanced tasks with ranges using the Excel JavaScript API. For the complete list of properties and methods that the `Range` object supports, see [Range Object (JavaScript API for Excel)](/javascript/api/excel/excel.range).

## Work with dates using the Moment-MSDate plug-in

The [Moment JavaScript library](https://momentjs.com/) provides a convenient way to use dates and timestamps. The [Moment-MSDate plug-in](https://www.npmjs.com/package/moment-msdate) converts the format of moments into one preferable for Excel. This is the same format the [NOW function](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) returns.

The following code shows how to set the range at **B4** to a moment's timestamp:

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var now = Date.now();
    var nowMoment = moment(now);
    var nowMS = nowMoment.toOADate();

    var dateRange = sheet.getRange("B4");
    dateRange.values = [[nowMS]];

    dateRange.numberFormat = [["[$-409]m/d/yy h:mm AM/PM;@"]];

    return context.sync();
}).catch(errorHandlerFunction);
```

It is a similar technique to get the date back out of the cell and convert it to a moment or other format, as demonstrated in the following code:

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var dateRange = sheet.getRange("B4");
    dateRange.load("values");

    return context.sync().then(function () {
        var nowMS = dateRange.values[0][0];

        // log the date as a moment
        var nowMoment = moment.fromOADate(nowMS);
        console.log(`get (moment): ${JSON.stringify(nowMoment)}`);

        // log the date as a UNIX-style timestamp
        var now = nowMoment.unix();
        console.log(`get (timestamp): ${now}`);
    });
}).catch(errorHandlerFunction);
```

Your add-in will have to format the ranges to display the dates in a more human-readable form. The example of `"[$-409]m/d/yy h:mm AM/PM;@"` displays a time like "12/3/18 3:57 PM". For more information about date and time number formats, please see the "Guidelines for date and time formats" in the [Review guidelines for customizing a number format](https://support.microsoft.com/office/c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5) article.

## Work with multiple ranges simultaneously

The [RangeAreas](/javascript/api/excel/excel.rangeareas) object lets your add-in perform operations on multiple ranges at once. These ranges may be contiguous, but do not have to be. `RangeAreas` are further discussed in the article [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).

## Find special cells within a range

The [Range.getSpecialCells](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-) and [Range.getSpecialCellsOrNullObject](/javascript/api/excel/excel.range#getspecialcellsornullobject-celltype--cellvaluetype-) methods find ranges based on the characteristics of their cells and the types of values of their cells. Both of these methods return `RangeAreas` objects. Here are the signatures of the methods from the TypeScript data types file:

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

The following example uses the `getSpecialCells` method to find all the cells with formulas. About this code, note:

- It limits the part of the sheet that needs to be searched by first calling `Worksheet.getUsedRange` and calling `getSpecialCells` for only that range.
- The `getSpecialCells` method returns a `RangeAreas` object, so all of the cells with formulas will be colored pink even if they are not all contiguous.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var formulaRanges = usedRange.getSpecialCells(Excel.SpecialCellType.formulas);
    formulaRanges.format.fill.color = "pink";

    return context.sync();
})
```

If no cells with the targeted characteristic exist in the range, `getSpecialCells` throws an **ItemNotFound** error. This diverts the flow of control to a `catch` block, if there is one. If there isn't a `catch` block, the error halts the method.

If you expect that cells with the targeted characteristic should always exist, you'll likely want your code to throw an error if those cells aren't there. If it's a valid scenario that there aren't any matching cells, your code should check for this possibility and handle it gracefully without throwing an error. You can achieve this behavior with the `getSpecialCellsOrNullObject` method and its returned `isNullObject` property. The following example uses this pattern. About this code, note:

- The `getSpecialCellsOrNullObject` method always returns a proxy object, so it is never `null` in the ordinary JavaScript sense. But if no matching cells are found, the `isNullObject` property of the object is set to `true`.
- It calls `context.sync` *before* it tests the `isNullObject` property. This is a requirement with all `*OrNullObject` methods and properties, because you always have to load and sync a property in order to read it. However, it is not necessary to *explicitly* load the `isNullObject` property. It is automatically loaded by the `context.sync` even if `load` is not called on the object. For more information, see [\*OrNullObject methods and properties](../develop/host-specific-api-model.md#ornullobject-methods-and-properties).
- You can test this code by first selecting a range that has no formula cells and running it. Then select a range that has at least one cell with a formula and run it again.

```js
Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    var formulaRanges = range.getSpecialCellsOrNullObject(Excel.SpecialCellType.formulas);
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

For simplicity, all other examples in this article use the `getSpecialCells` method instead of  `getSpecialCellsOrNullObject`.

### Narrow the target cells with cell value types

The `Range.getSpecialCells()` and `Range.getSpecialCellsOrNullObject()` methods accept an optional second parameter used to further narrow down the targeted cells. This second parameter is an `Excel.SpecialCellValueType` you use to specify that you only want cells that contain certain types of values.

> [!NOTE]
> The `Excel.SpecialCellValueType` parameter can only be used if the `Excel.SpecialCellType` is `Excel.SpecialCellType.formulas` or `Excel.SpecialCellType.constants`.

#### Test for a single cell value type

The `Excel.SpecialCellValueType` enum has these four basic types (in addition to the other combined values described later in this section):

- `Excel.SpecialCellValueType.errors`
- `Excel.SpecialCellValueType.logical` (which means boolean)
- `Excel.SpecialCellValueType.numbers`
- `Excel.SpecialCellValueType.text`

The following example finds special cells that are numerical constants and colors those cells pink. About this code, note:

- It will only highlight cells that have a literal number value. It will not highlight cells that have a formula (even if the result is a number) or a boolean, text, or error state cells.
- To test the code, be sure the worksheet has some cells with literal number values, some with other kinds of literal values, and some with formulas.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var constantNumberRanges = usedRange.getSpecialCells(
        Excel.SpecialCellType.constants,
        Excel.SpecialCellValueType.numbers);
    constantNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

#### Test for multiple cell value types

Sometimes you need to operate on more than one cell value type, such as all text-valued and all boolean-valued (`Excel.SpecialCellValueType.logical`) cells. The `Excel.SpecialCellValueType` enum has values with combined types. For example, `Excel.SpecialCellValueType.logicalText` targets all boolean and all text-valued cells. `Excel.SpecialCellValueType.all` is the default value, which does not limit the cell value types returned. The following example colors all cells with formulas that produce number or boolean value.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var formulaLogicalNumberRanges = usedRange.getSpecialCells(
        Excel.SpecialCellType.formulas,
        Excel.SpecialCellValueType.logicalNumbers);
    formulaLogicalNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

## Cut, copy, and paste

### Copy and paste

The [Range.copyFrom](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-) method replicates the **Copy** and **Paste** actions of the Excel UI. The range object that `copyFrom` is called on is the destination. The source to be copied is passed as a range or a string address representing a range.

The following code sample copies the data from **A1:E1** into the range starting at **G1** (which ends up pasting into **G1:K1**).

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy everything from "A1:E1" into "G1" and the cells afterwards ("G1:K1")
    sheet.getRange("G1").copyFrom("A1:E1");
    return context.sync();
}).catch(errorHandlerFunction);
```

`Range.copyFrom` has three optional parameters.

```TypeScript
copyFrom(sourceRange: Range | RangeAreas | string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean): void;
```

`copyType` specifies what data gets copied from the source to the destination.

- `Excel.RangeCopyType.formulas` transfers the formulas in the source cells and preserves the relative positioning of those formulas' ranges. Any non-formula entries are copied as-is.
- `Excel.RangeCopyType.values` copies the data values and, in the case of formulas, the result of the formula.
- `Excel.RangeCopyType.formats` copies the formatting of the range, including font, color, and other format settings, but no values.
- `Excel.RangeCopyType.all` (the default option) copies both data and formatting, preserving cells' formulas if found.

`skipBlanks` sets whether blank cells are copied into the destination. When true, `copyFrom` skips blank cells in the source range.
Skipped cells will not overwrite the existing data of their corresponding cells in the destination range. The default is false.

`transpose` determines whether or not the data is transposed, meaning its rows and columns are switched, into the source location.
A transposed range is flipped along the main diagonal, so rows **1**, **2**, and **3** will become columns **A**, **B**, and **C**.

The following code sample and images demonstrate this behavior in a simple scenario.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy a range, omitting the blank cells so existing data is not overwritten in those cells
    sheet.getRange("D1").copyFrom("A1:C1",
        Excel.RangeCopyType.all,
        true, // skipBlanks
        false); // transpose
    // copy a range, including the blank cells which will overwrite existing data in the target cells
    sheet.getRange("D2").copyFrom("A2:C2",
        Excel.RangeCopyType.all,
        false, // skipBlanks
        false); // transpose
    return context.sync();
}).catch(errorHandlerFunction);
```

*Before the preceding function has been run.*

![Data in Excel before range's copy method has been run](../images/excel-range-copyfrom-skipblanks-before.png)

*After the preceding function has been run.*

![Data in Excel after range's copy method has been run](../images/excel-range-copyfrom-skipblanks-after.png)

### Cut and paste (move) cells

The [Range.moveTo](/javascript/api/excel/excel.range#moveto-destinationrange-) method moves cells to a new location in the workbook. This cell movement behavior works the same as when cells are moved by [dragging the range border](https://support.office.com/article/Move-or-copy-cells-and-cell-contents-803d65eb-6a3e-4534-8c6f-ff12d1c4139e) or when taking the **Cut** and **Paste** actions. Both the formatting and values of the range are moved to the location specified as the `destinationRange` parameter.

The following code sample shows a range being moved with the `Range.moveTo` method. Note that if the destination range is smaller than the source, it will be expanded to encompass the source content.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.getRange("F1").values = [["Moved Range"]];

    // Move the cells "A1:E1" to "G1" (which fills the range "G1:K1").
    sheet.getRange("A1:E1").moveTo("G1");
    return context.sync();
});
```

## Remove duplicates

The [Range.removeDuplicates](/javascript/api/excel/excel.range#removeduplicates-columns--includesheader-) method removes rows with duplicate entries in the specified columns. The method goes through each row in the range from the lowest-valued index to the highest-valued index in the range (from top to bottom). A row is deleted if a value in its specified column or columns appeared earlier in the range. Rows in the range below the deleted row are shifted up. `removeDuplicates` does not affect the position of cells outside of the range.

`removeDuplicates` takes in a `number[]` representing the column indices which are checked for duplicates. This array is zero-based and relative to the range, not the worksheet. The method also takes in a boolean parameter that specifies whether the first row is a header. When **true**, the top row is ignored when considering duplicates. The `removeDuplicates` method returns a `RemoveDuplicatesResult` object that specifies the number of rows removed and the number of unique rows remaining.

When using a range's `removeDuplicates` method, keep the following in mind:

- `removeDuplicates` considers cell values, not function results. If two different functions evaluate to the same result, the cell values are not considered duplicates.
- Empty cells are not ignored by `removeDuplicates`. The value of an empty cell is treated like any other value. This means empty rows contained within in the range will be included in the `RemoveDuplicatesResult`.

The following sample shows the removal of entries with duplicate values in the first column.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:D11");

    var deleteResult = range.removeDuplicates([0],true);
    deleteResult.load();

    return context.sync().then(function () {
        console.log(deleteResult.removed + " entries with duplicate names removed.");
        console.log(deleteResult.uniqueRemaining + " entries with unique names remain in the range.");
    });
}).catch(errorHandlerFunction);
```

*Before the preceding function has been run.*

![Data in Excel before range's remove duplicates method has been run](../images/excel-ranges-remove-duplicates-before.png)

*After the preceding function has been run.*

![Data in Excel after range's remove duplicates method has been run](../images/excel-ranges-remove-duplicates-after.png)

## Group data for an outline

Rows or columns of a range can be grouped together to create an [outline](https://support.office.com/article/Outline-group-data-in-a-worksheet-08CE98C4-0063-4D42-8AC7-8278C49E9AFF). These groups can be collapsed and expanded to hide and show the corresponding cells. This makes quick analysis of top-line data easier. Use [Range.group](/javascript/api/excel/excel.range#group-groupoption-) to make these outline groups.

An outline can have a hierarchy, where smaller groups are nested under larger groups. This allows the outline to be viewed at different levels. Changing the visible outline level can be done programmatically through the [Worksheet.showOutlineLevels](/javascript/api/excel/excel.worksheet#showoutlinelevels-rowlevels--columnlevels-) method. Note that Excel only supports eight levels of outline groups.

The following code sample shows how to create an outline with two levels of groups for both the rows and columns. The subsequent image shows the groupings of that outline. Note that in the code sample, the ranges being grouped do not include the row or column of the outline control (the "Totals" for this example). A group defines what will be collapsed, not the row or column with the control.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    // Group the larger, main level. Note that the outline controls
    // will be on row 10, meaning 4-9 will collapse and expand.
    sheet.getRange("4:9").group(Excel.GroupOption.byRows);

    // Group the smaller, sublevels. Note that the outline controls
    // will be on rows 6 and 9, meaning 4-5 and 7-8 will collapse and expand.
    sheet.getRange("4:5").group(Excel.GroupOption.byRows);
    sheet.getRange("7:8").group(Excel.GroupOption.byRows);

    // Group the larger, main level. Note that the outline controls
    // will be on column R, meaning C-Q will collapse and expand.
    sheet.getRange("C:Q").group(Excel.GroupOption.byColumns);

    // Group the smaller, sublevels. Note that the outline controls
    // will be on columns G, L, and R, meaning C-F, H-K, and M-P will collapse and expand.
    sheet.getRange("C:F").group(Excel.GroupOption.byColumns);
    sheet.getRange("H:K").group(Excel.GroupOption.byColumns);
    sheet.getRange("M:P").group(Excel.GroupOption.byColumns);
    return context.sync();
}).catch(errorHandlerFunction);

```

![A range with a two-level, two-dimension outline](../images/excel-outline.png)

To ungroup a row or column group, use the [Range.ungroup](/javascript/api/excel/excel.range#ungroup-groupoption-) method. This removes the outermost level from the outline. If multiple groups of the same row or column type are at the same level within the specified range, all of those groups are ungrouped.

## See also

- [Work with ranges using the Excel JavaScript API](excel-add-ins-ranges.md)
- [Fundamental programming concepts with the Excel JavaScript API](excel-add-ins-core-concepts.md)
- [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md)
