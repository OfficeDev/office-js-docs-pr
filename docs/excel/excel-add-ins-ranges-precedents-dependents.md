---
title: Work with formula precedents and dependents using the Excel JavaScript API
description: 'Learn how to use the Excel JavaScript API to retrieve formula precedents and dependents.' 
ms.date: 05/28/2021
ms.prod: excel
localization_priority: Normal
---

# Get formula precedents and dependents using the Excel JavaScript API

This article provides code samples that retrieve formula precedents and formula dependents using the Excel JavaScript API. For the complete list of properties and methods that the `Range` object supports, see [Range Object (JavaScript API for Excel)](/javascript/api/excel/excel.range).

An Excel formula often refers to other cells. When a cell provides data to a formula, it is known as a formula "precedent". When a cell contains a formula that refers to other cells, it is known as a formula "dependent". To learn more about Excel features related to relationships between cells, see [Display the relationships between formulas and cells](https://support.microsoft.com/office/display-the-relationships-between-formulas-and-cells-a59bef2b-3701-46bf-8ff1-d3518771d507).

## Get formula precedents

With [Range.getDirectPrecedents](/javascript/api/excel/excel.range#getdirectprecedents--), your add-in can locate a formula's direct precedent cells. `Range.getDirectPrecedents` returns a `WorkbookRangeAreas` object. This object contains the addresses of all the precedents in the workbook. It has a separate `RangeAreas` object for each worksheet containing at least one formula precedent. For more information on working with the `RangeAreas` object, see [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).

In the Excel UI, the **Trace Precedents** button draws an arrow from precedent cells to the selected formula. Unlike the Excel UI button, the `getDirectPrecedents` method does not draw arrows.

> [!IMPORTANT]
> The `getDirectPrecedents` method can't retrieve precedent cells across workbooks.

The following code sample gets the direct precedents for the active range and then changes the background color of those precedent cells to yellow.

> [!NOTE]
> The active range must contain a formula that references other cells in the same workbook for the highlighting to work properly.

```js
Excel.run(function (context) {
    // Precedents are cells that provide data to the selected formula.
    var range = context.workbook.getActiveCell();
    var directPrecedents = range.getDirectPrecedents();
    range.load("address");
    directPrecedents.areas.load("address");
    
    return context.sync()
        .then(function () {
            console.log(`Direct precedent cells of ${range.address}:`);

            // Use the direct precedents API to loop through precedents of the active cell.
            for (var i = 0; i < directPrecedents.areas.items.length; i++) {
              // Highlight and print out the address of each precedent cell.
              directPrecedents.areas.items[i].format.fill.color = "Yellow";
              console.log(`  ${directPrecedents.areas.items[i].address}`);
            }
        })
        .then(context.sync);
}).catch(errorHandlerFunction);
```

## Get formula dependents (preview)

> [!NOTE]
> The `Range.getDirectDependents` method is currently only available in public preview. [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]
> 

With [Range.getDirectDependents](/javascript/api/excel/excel.range#getDirectDependents__), your add-in can locate a formula's direct dependent cells. Like `Range.getDirectPrecedents`, `Range.getDirectDependents` also returns a `WorkbookRangeAreas` object. This object contains the addresses of all the dependents in the workbook. It has a separate `RangeAreas` object for each worksheet containing at least one formula dependent. For more information on working with the `RangeAreas` object, see [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).

In the Excel UI, the **Trace Dependents** button draws an arrow from dependent cells to the selected formula. Unlike the Excel UI button, the `getDirectDependents` method does not draw arrows.

> [!IMPORTANT]
> The `getDirectDependents` method can't retrieve dependent cells across workbooks.

The following code sample gets the direct dependents for the active range and then changes the background color of those dependent cells to yellow.

> [!NOTE]
> The active range must contain a formula that references other cells in the same workbook for the highlighting to work properly.

```js
Excel.run(function (context) {
    // Direct dependents are cells that contain formulas that refer to other cells.
    let range = context.workbook.getActiveCell();
    let directDependents = range.getDirectDependents();
    range.load("address");
    directDependents.areas.load("address");
    
    return context.sync()
        .then(function () {
            console.log(`Direct dependent cells of ${range.address}:`);
    
            // Use the direct dependents API to loop through direct dependents of the active cell.
            for (let i = 0; i < directDependents.areas.items.length; i++) {
              // Highlight and print the address of each dependent cell.
              directDependents.areas.items[i].format.fill.color = "Yellow";
              console.log(`  ${directDependents.areas.items[i].address}`);
            }
        })
        .then(context.sync);
}).catch(errorHandlerFunction);
```

## See also

- [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md)
- [Work with cells using the Excel JavaScript API](excel-add-ins-cells.md)
- [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md)
