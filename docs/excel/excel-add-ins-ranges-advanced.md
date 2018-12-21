---
title: Work with ranges using the Excel JavaScript API (advanced)
description: ''
ms.date: 12/18/2018
---

# Work with ranges using the Excel JavaScript API (advanced)

This article builds upon information in [Work with ranges using the Excel JavaScript API (fundamental)](excel-add-ins-ranges.md) by providing code samples that show how to perform more advanced tasks with ranges using the Excel JavaScript API. For the complete list of properties and methods that the **Range** object supports, see [Range Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.range).

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

Your add-in will have to format the ranges to display the dates in a more human-readable form. The example of `"[$-409]m/d/yy h:mm AM/PM;@"` displays a time like "12/3/18 3:57 PM". For more information about date and time number formats, please see the "Guidelines for date and time formats" in the [Review guidelines for customizing a number format](https://support.office.com/article/review-guidelines-for-customizing-a-number-format-c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5) article.

## Work with multiple ranges simultaneously (preview)

> [!NOTE]
> The `RangeAreas` object is currently available only in public preview (beta). To use this feature, you must use the beta library of the Office.js CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.
> If you are using TypeScript or your code editor uses TypeScript type definition files for IntelliSense, use https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.

The `RangeAreas` object lets your add-in perform operations on multiple ranges at once. These ranges may be contiguous, but do not have to be. `RangeAreas` are further discussed in the article [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).

## Copy and paste (preview)

> [!NOTE]
> The `Range.copyFrom` function is currently available only in public preview (beta). To use this feature, you must use the beta library of the Office.js CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.
> If you are using TypeScript or your code editor uses TypeScript type definition files for IntelliSense, use https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.

Range’s `copyFrom` function replicates the copy-and-paste behavior of the Excel UI. The range object that `copyFrom` is called on is the destination.
The source to be copied is passed as a range or a string address representing a range. 
The following code sample copies the data from **A1:E1** into the range starting at **G1** (which ends up pasting into **G1:K1**).

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy a range starting at a single cell destination
    sheet.getRange("G1").copyFrom("A1:E1");
    return context.sync();
}).catch(errorHandlerFunction);
```

`Range.copyFrom` has three optional parameters.

```TypeScript
copyFrom(sourceRange: Range | string, copyType?: "All" | "Formulas" | "Values" | "Formats", skipBlanks?: boolean, transpose?: boolean): void;
```

`copyType` specifies what data gets copied from the source to the destination.

- `"Formulas"` transfers the formulas in the source cells and preserves the relative positioning of those formulas’ ranges. Any non-formula entries are copied as-is.
- `"Values"` copies the data values and, in the case of formulas, the result of the formula.
- `"Formats"` copies the formatting of the range, including font, color, and other format settings, but no values.
- `"All"` (the default option) copies both data and formatting, preserving cells’ formulas if found.

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

![Data in Excel before range’s copy method has been run](../images/excel-range-copyfrom-skipblanks-before.png)

*After the preceding function has been run.*

![Data in Excel after range’s copy method has been run](../images/excel-range-copyfrom-skipblanks-after.png)

## Remove duplicates (preview)

> [!NOTE]
> The Range object's `removeDuplicates` function is currently available only in public preview (beta). To use this feature, you must use the beta library of the Office.js CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.
> If you are using TypeScript or your code editor uses TypeScript type definition files for IntelliSense, use https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.

The Range object's `removeDuplicates` function removes rows with duplicate entries in the specified columns. The function goes through each row in the range from the lowest-valued index to the highest-valued index in the range (from top to bottom). A row is deleted if a value in its specified column or columns appeared earlier in the range. Rows in the range below the deleted row are shifted up. `removeDuplicates` does not affect the position of cells outside of the range.

`removeDuplicates` takes in a `number[]` representing the column indices which are checked for duplicates. This array is zero-based and relative to the range, not the worksheet. The function also takes in a boolean parameter that specifies whether the first row is a header. When **true**, the top row is ignored when considering duplicates. The `removeDuplicates` function returns a `RemoveDuplicatesResult` object that specifies the number of rows removed and the number of unique rows remaining.

When using a range's `removeDuplicates` function, keep the following in mind:

- `removeDuplicates` considers cell values, not function results. If two different functions evaluate to the same result, the cell values are not considered duplicates.
- Empty cells are not ignored by `removeDuplicates`. The value of an empty cell is treated like any other value. This means empty rows contained within in the range will be included in the `RemoveDuplicatesResult`.

The following sample shows the removal of entries with duplicate values in the first column.

```js
Excel.run(async (context) => {
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

![Data in Excel before range’s remove duplicates method has been run](../images/excel-ranges-remove-duplicates-before.png)

*After the preceding function has been run.*

![Data in Excel after range’s remove duplicates method has been run](../images/excel-ranges-remove-duplicates-after.png)

## See also

- [Work with ranges using the Excel JavaScript API](excel-add-ins-ranges.md)
- [Fundamental programming concepts with the Excel JavaScript API](excel-add-ins-core-concepts.md)