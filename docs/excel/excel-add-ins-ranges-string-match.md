---
title: Find strings in an Excel worksheet using the JavaScript API
description: Learn how to find the first match, find all matches, handle missing results, and control partial and case-sensitive searches using the Excel JavaScript API.
ms.date: 03/26/2026
ms.localizationpriority: medium
---

# Find strings in an Excel worksheet using the JavaScript API

Searching for text is one of the most common operations in Excel. The Excel JavaScript API provides two ways to find strings programmatically.

- **Find the first match**: Use `Range.find` to locate the first cell in a range that contains your search text.
- **Find every match**: Use `Worksheet.findAll` to locate every cell on a worksheet that contains your search text and act on the results as a group.

Both methods accept a [SearchCriteria](/javascript/api/excel/excel.searchcriteria) object that you can use to control whether the search matches whole cell values or partial text, whether it's case-sensitive, and (for `Range.find`) which direction it searches.

This article walks through common search scenarios by using these methods and shows how to handle cases where no match exists.

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## Find the first matching cell

Use `Range.find` when you only need the first cell that matches your search text. It returns a `Range` representing that cell. You pass the search string as the first argument and an optional `SearchCriteria` object as the second.

The following code sample searches the used range of the active worksheet for the first cell whose entire value equals **Apples**. The `completeMatch` option ensures that a cell containing "Green Apples" isn't returned. After the search, the cell address is logged to the console.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let searchRange = sheet.getUsedRange();

    // Find the first cell that exactly equals "Apples".
    let foundRange = searchRange.find("Apples", {
        completeMatch: true, // Match the whole cell value.
        matchCase: false, // Case-insensitive search.
        searchDirection: Excel.SearchDirection.forward // Search from the start of the range.
    });

    foundRange.load("address");
    await context.sync();

    // If the worksheet contains "Apples" in cell B2, this logs "B2".
    console.log(foundRange.address);
});
```

> [!NOTE]
> When you call `find` on a range that represents a single cell, the entire worksheet is searched. The search begins at that cell and proceeds in the direction specified by `SearchCriteria.searchDirection`, wrapping around the ends of the worksheet if needed.

`find` throws an `ItemNotFound` error if the string doesn't exist in the range. See [Handle missing matches](#handle-missing-matches) later in this article.

## Find every matching cell on a worksheet

Use [`Worksheet.findAll`](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-findall-member(1)) when you need to locate *every* cell that contains the search text and act on them together. It returns a `RangeAreas` object, which is a collection of `Range` objects that you can format or edit in a single operation.

The following code sample finds all cells on the worksheet that contain the value **Complete** and colors them green.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();

    let foundRanges = sheet.findAll("Complete", {
        completeMatch: true, // Match the whole cell value.
        matchCase: false // Case-insensitive search.
    });

    await context.sync();
    foundRanges.format.fill.color = "green";
});
```

Like `Range.find`, `Worksheet.findAll` throws an `ItemNotFound` error when there are no matches. See the next section for how to handle this error gracefully.

## Handle missing matches

Both `find` and `findAll` throw an `ItemNotFound` error when the search text isn't found. If the text might not exist, use the `*OrNullObject` variant of each method so your code can handle both outcomes without an error.

- `Range.findOrNullObject`: returns a proxy object whose `isNullObject` property is `true` when no match is found.
- `Worksheet.findAllOrNullObject`: works the same way, returning a `RangeAreas` proxy with `isNullObject` set to `true`.

For more information about the pattern, see [*OrNullObject methods and properties](../develop/application-specific-api-model.md#ornullobject-methods-and-properties).

The following code sample uses `findOrNullObject` to search for a value that might be missing.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let searchRange = sheet.getUsedRange();

    let foundRange = searchRange.findOrNullObject("Mangoes", {
        completeMatch: true,
        matchCase: false,
        searchDirection: Excel.SearchDirection.forward
    });

    foundRange.load(["address", "isNullObject"]);
    await context.sync();

    if (foundRange.isNullObject) {
        console.log("No cell contains 'Mangoes'.");
    } else {
        console.log(`Found 'Mangoes' at ${foundRange.address}.`);
    }
});
```

## Find a partial match

By default, `find` and `findAll` perform a partial match. Setting `completeMatch` to `false` (or omitting it) returns cells that *contain* the search string anywhere in their value.

The following example finds the first cell that contains the text **fruit** as part of a longer string, such as "Grapefruit" or "Kiwifruit".

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let searchRange = sheet.getUsedRange();

    // Partial match: returns a cell containing "Grapefruit", "Kiwifruit", etc.
    let foundRange = searchRange.find("fruit", {
        completeMatch: false, // Match any cell that contains the substring.
        matchCase: false,
        searchDirection: Excel.SearchDirection.forward
    });

    foundRange.load(["address", "values"]);
    await context.sync();

    console.log(`${foundRange.values[0][0]} found at ${foundRange.address}`);
});
```

## Run a case-sensitive search

Set `matchCase` to `true` when the casing of the search text matters. The following example finds the first cell that contains **GDP** without matching cells that contain "gdp" or "Gdp".

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let searchRange = sheet.getUsedRange();

    let foundRange = searchRange.find("GDP", {
        completeMatch: false,
        matchCase: true, // Only match the exact casing "GDP".
        searchDirection: Excel.SearchDirection.forward
    });

    foundRange.load("address");
    await context.sync();

    console.log(foundRange.address);
});
```

## See also

- [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md)
- [Work with cells using the Excel JavaScript API](excel-add-ins-cells.md)
- [Find special cells within a range using the Excel JavaScript API](excel-add-ins-ranges-special-cells.md)
- [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md)
- [Set and get range values, text, or formulas using the Excel JavaScript API](excel-add-ins-ranges-set-get-values.md)
