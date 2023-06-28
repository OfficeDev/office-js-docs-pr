---
title: Find special cells within a range using the Excel JavaScript API
description: Learn how to use the Excel JavaScript API to find special cells, such as cells with formulas, errors, or numbers.
ms.date: 02/17/2022
ms.localizationpriority: medium
---

# Find special cells within a range using the Excel JavaScript API

This article provides code samples that find special cells within a range using the Excel JavaScript API. For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).

## Find ranges with special cells

The [Range.getSpecialCells](/javascript/api/excel/excel.range#excel-excel-range-getspecialcells-member(1)) and [Range.getSpecialCellsOrNullObject](/javascript/api/excel/excel.range#excel-excel-range-getspecialcellsornullobject-member(1)) methods find ranges based on the characteristics of their cells and the types of values of their cells. Both of these methods return `RangeAreas` objects. Here are the signatures of the methods from the TypeScript data types file:

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

The following code sample uses the `getSpecialCells` method to find all the cells with formulas. About this code, note:

- It limits the part of the sheet that needs to be searched by first calling `Worksheet.getUsedRange` and calling `getSpecialCells` for only that range.
- The `getSpecialCells` method returns a `RangeAreas` object, so all of the cells with formulas will be colored pink even if they are not all contiguous.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let usedRange = sheet.getUsedRange();
    let formulaRanges = usedRange.getSpecialCells(Excel.SpecialCellType.formulas);
    formulaRanges.format.fill.color = "pink";

    await context.sync();
});
```

If no cells with the targeted characteristic exist in the range, `getSpecialCells` throws an **ItemNotFound** error. This diverts the flow of control to a `catch` block, if there is one. If there isn't a `catch` block, the error halts the method.

If you expect that cells with the targeted characteristic should always exist, you'll likely want your code to throw an error if those cells aren't there. If it's a valid scenario that there aren't any matching cells, your code should check for this possibility and handle it gracefully without throwing an error. You can achieve this behavior with the `getSpecialCellsOrNullObject` method and its returned `isNullObject` property. The following code sample uses this pattern. About this code, note:

- The `getSpecialCellsOrNullObject` method always returns a proxy object, so it's never `null` in the ordinary JavaScript sense. But if no matching cells are found, the `isNullObject` property of the object is set to `true`.
- It calls `context.sync` *before* it tests the `isNullObject` property. This is a requirement with all `*OrNullObject` methods and properties, because you always have to load and sync a property in order to read it. However, it's not necessary to *explicitly* load the `isNullObject` property. It's automatically loaded by the `context.sync` even if `load` is not called on the object. For more information, see [\*OrNullObject methods and properties](../develop/application-specific-api-model.md#ornullobject-methods-and-properties).
- You can test this code by first selecting a range that has no formula cells and running it. Then select a range that has at least one cell with a formula and run it again.

```js
await Excel.run(async (context) => {
    let range = context.workbook.getSelectedRange();
    let formulaRanges = range.getSpecialCellsOrNullObject(Excel.SpecialCellType.formulas);
    await context.sync();
        
    if (formulaRanges.isNullObject) {
        console.log("No cells have formulas");
    }
    else {
        formulaRanges.format.fill.color = "pink";
    }
    
    await context.sync();
});
```

For simplicity, all other code samples in this article use the `getSpecialCells` method instead of  `getSpecialCellsOrNullObject`.

## Narrow the target cells with cell value types

The `Range.getSpecialCells()` and `Range.getSpecialCellsOrNullObject()` methods accept an optional second parameter used to further narrow down the targeted cells. This second parameter is an `Excel.SpecialCellValueType` you use to specify that you only want cells that contain certain types of values.

> [!NOTE]
> The `Excel.SpecialCellValueType` parameter can only be used if the `Excel.SpecialCellType` is `Excel.SpecialCellType.formulas` or `Excel.SpecialCellType.constants`.

### Test for a single cell value type

The `Excel.SpecialCellValueType` enum has these four basic types (in addition to the other combined values described later in this section):

- `Excel.SpecialCellValueType.errors`
- `Excel.SpecialCellValueType.logical` (which means boolean)
- `Excel.SpecialCellValueType.numbers`
- `Excel.SpecialCellValueType.text`

The following code sample finds special cells that are numerical constants and colors those cells pink. About this code, note:

- It only highlights cells that have a literal number value. It won't highlight cells that have a formula (even if the result is a number) or a boolean, text, or error state cells.
- To test the code, be sure the worksheet has some cells with literal number values, some with other kinds of literal values, and some with formulas.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let usedRange = sheet.getUsedRange();
    let constantNumberRanges = usedRange.getSpecialCells(
        Excel.SpecialCellType.constants,
        Excel.SpecialCellValueType.numbers);
    constantNumberRanges.format.fill.color = "pink";

    await context.sync();
});
```

### Test for multiple cell value types

Sometimes you need to operate on more than one cell value type, such as all text-valued and all boolean-valued (`Excel.SpecialCellValueType.logical`) cells. The `Excel.SpecialCellValueType` enum has values with combined types. For example, `Excel.SpecialCellValueType.logicalText` targets all boolean and all text-valued cells. `Excel.SpecialCellValueType.all` is the default value, which does not limit the cell value types returned. The following code sample colors all cells with formulas that produce number or boolean value.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let usedRange = sheet.getUsedRange();
    let formulaLogicalNumberRanges = usedRange.getSpecialCells(
        Excel.SpecialCellType.formulas,
        Excel.SpecialCellValueType.logicalNumbers);
    formulaLogicalNumberRanges.format.fill.color = "pink";

    await context.sync();
});
```

## See also

- [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md)
- [Work with cells using the Excel JavaScript API](excel-add-ins-cells.md)
- [Find a string using the Excel JavaScript API](excel-add-ins-ranges-string-match.md)
- [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md)
