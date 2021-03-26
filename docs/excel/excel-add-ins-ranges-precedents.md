---
title: Work with formula precedents using the Excel JavaScript API
description: '' 
ms.date: 03/26/2021 
localization_priority: Normal
---

# Work with ranges using the Excel JavaScript API (advanced)

This article builds upon information in [Work with ranges using the Excel JavaScript API (fundamental)](excel-add-ins-ranges.md) by providing code samples that show how to perform more advanced tasks with ranges using the Excel JavaScript API. For the complete list of properties and methods that the `Range` object supports, see [Range Object (JavaScript API for Excel)](/javascript/api/excel/excel.range).

## Get formula precedents

An Excel formula often refers to other cells. When a cell provides data to a formula, it is known as a formula "precedent". To learn more about Excel features related to relationships between cells, see the [Display the relationships between formulas and cells](https://support.microsoft.com/office/display-the-relationships-between-formulas-and-cells-a59bef2b-3701-46bf-8ff1-d3518771d507) article. 

With [Range.getDirectPrecedents](/javascript/api/excel/excel.range#getdirectprecedents--), your add-in can locate a formula's direct precedent cells. `Range.getDirectPrecedents` returns a `WorkbookRangeAreas` object. This object contains the addresses of all the precedents in the workbook. It has a separate `RangeAreas` object for each worksheet containing at least one formula precedent. See [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md) for more information on working with the `RangeAreas` object.

In the Excel UI, the **Trace Precedents** button draws an arrow from precedent cells to the selected formula. Unlike the Excel UI button, the `getDirectPrecedents` method does not draw arrows. 

> [!IMPORTANT]
> The `getDirectPrecedents` method can't retrieve precedent cells across workbooks. 

The following sample gets the direct precedents for the active range and then changes the background color of those precedent cells to yellow. 

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

## See also

- [Work with ranges using the Excel JavaScript API](excel-add-ins-ranges.md)
- [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md)
- [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md)
