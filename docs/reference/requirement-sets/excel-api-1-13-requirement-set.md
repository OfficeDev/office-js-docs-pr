---
title: Excel JavaScript API requirement set 1.13
description: 'Details about the ExcelApi 1.13 requirement set.'
ms.date: 07/09/2021
ms.prod: excel
ms.localizationpriority: medium
---

# What's new in Excel JavaScript API 1.13

The ExcelApi 1.13 added a method to insert worksheets into a workbook from a Base64-encoded string and an event to detect workbook activation. It also increased support for formulas in ranges by adding APIs to track changes to formulas and locate a formula's direct dependent cells. Additionally, it expanded PivotTable support by adding PivotLayout APIs for alt text, style, and empty cell management.

| Feature area | Description | Relevant objects |
|:--- |:--- |:--- |
| [Formula changed events](../../excel/excel-add-ins-worksheets.md#detect-formula-changes) | Track changes to formulas, including the source and type of event that caused a change. | [Worksheet.onFormulaChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onformulachanged-member)|
| [Formula dependents](../../excel/excel-add-ins-ranges-precedents-dependents.md#get-the-direct-dependents-of-a-formula) | Locate the direct dependent cells of a formula. | [Range.getDirectDependents](/javascript/api/excel/excel.range#excel-excel-range-getdirectdependents-member(1)) |
| [Insert worksheets](../../excel/excel-add-ins-workbooks.md#insert-a-copy-of-an-existing-workbook-into-the-current-one) | Insert worksheets from another workbook into the current workbook as a Base64-encoded string. | [Workbook.insertWorksheetsFromBase64](/javascript/api/excel/excel.workbook#excel-excel-workbook-insertworksheetsfrombase64-member(1)) |
| [PivotTable PivotLayout](../../excel/excel-add-ins-pivottables.md#other-pivotlayout-functions) | An expansion of the PivotLayout class, including new support for alt text and empty cell management. | [PivotLayout](/javascript/api/excel/excel.pivotlayout) |

## API list

The following table lists the APIs in Excel JavaScript API requirement set 1.13. To view API reference documentation for all APIs supported by Excel JavaScript API requirement set 1.13 or earlier, see [Excel APIs in requirement set 1.13 or earlier](/javascript/api/excel?view=excel-js-1.13&preserve-view=true).

| Class | Fields | Description |
|:---|:---|:---|
|[FormulaChangedEventDetail](/javascript/api/excel/excel.formulachangedeventdetail)|[cellAddress](/javascript/api/excel/excel.formulachangedeventdetail#excel-excel-formulachangedeventdetail-celladdress-member)|The address of the cell that contains the changed formula.|
||[previousFormula](/javascript/api/excel/excel.formulachangedeventdetail#excel-excel-formulachangedeventdetail-previousformula-member)|Represents the previous formula, before it was changed.|
|[InsertWorksheetOptions](/javascript/api/excel/excel.insertworksheetoptions)|[positionType](/javascript/api/excel/excel.insertworksheetoptions#excel-excel-insertworksheetoptions-positiontype-member)|The insert position, in the current workbook, of the new worksheets.|
||[relativeTo](/javascript/api/excel/excel.insertworksheetoptions#excel-excel-insertworksheetoptions-relativeto-member)|The worksheet in the current workbook that is referenced for the `WorksheetPositionType` parameter.|
||[sheetNamesToInsert](/javascript/api/excel/excel.insertworksheetoptions#excel-excel-insertworksheetoptions-sheetnamestoinsert-member)|The names of individual worksheets to insert.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[altTextDescription](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-alttextdescription-member)|The alt text description of the PivotTable.|
||[altTextTitle](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-alttexttitle-member)|The alt text title of the PivotTable.|
||[displayBlankLineAfterEachItem(display: boolean)](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-displayblanklineaftereachitem-member(1))|Sets whether or not to display a blank line after each item.|
||[emptyCellText](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-emptycelltext-member)|The text that is automatically filled into any empty cell in the PivotTable if `fillEmptyCells == true`.|
||[fillEmptyCells](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-fillemptycells-member)|Specifies whether empty cells in the PivotTable should be populated with the `emptyCellText`.|
||[repeatAllItemLabels(repeatLabels: boolean)](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-repeatallitemlabels-member(1))|Sets the "repeat all item labels" setting across all fields in the PivotTable.|
||[showFieldHeaders](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-showfieldheaders-member)|Specifies whether the PivotTable displays field headers (field captions and filter drop-downs).|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[refreshOnOpen](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-refreshonopen-member)|Specifies whether the PivotTable refreshes when the workbook opens.|
|[Range](/javascript/api/excel/excel.range)|[getDirectDependents()](/javascript/api/excel/excel.range#excel-excel-range-getdirectdependents-member(1))|Returns a `WorkbookRangeAreas` object that represents the range containing all the direct dependents of a cell in the same worksheet or in multiple worksheets.|
||[getExtendedRange(direction: Excel.KeyboardDirection, activeCell?: Range \| string)](/javascript/api/excel/excel.range#excel-excel-range-getextendedrange-member(1))|Returns a range object that includes the current range and up to the edge of the range, based on the provided direction.|
||[getMergedAreasOrNullObject()](/javascript/api/excel/excel.range#excel-excel-range-getmergedareasornullobject-member(1))|Returns a RangeAreas object that represents the merged areas in this range.|
||[getRangeEdge(direction: Excel.KeyboardDirection, activeCell?: Range \| string)](/javascript/api/excel/excel.range#excel-excel-range-getrangeedge-member(1))|Returns a range object that is the edge cell of the data region that corresponds to the provided direction.|
|[Table](/javascript/api/excel/excel.table)|[resize(newRange: Range \| string)](/javascript/api/excel/excel.table#excel-excel-table-resize-member(1))|Resize the table to the new range.|
|[Workbook](/javascript/api/excel/excel.workbook)|[insertWorksheetsFromBase64(base64File: string, options?: Excel.InsertWorksheetOptions)](/javascript/api/excel/excel.workbook#excel-excel-workbook-insertworksheetsfrombase64-member(1))|Inserts the specified worksheets from a source workbook into the current workbook.|
||[onActivated](/javascript/api/excel/excel.workbook#excel-excel-workbook-onactivated-member)|Occurs when the workbook is activated.|
|[WorkbookActivatedEventArgs](/javascript/api/excel/excel.workbookactivatedeventargs)|[type](/javascript/api/excel/excel.workbookactivatedeventargs#excel-excel-workbookactivatedeventargs-type-member)|Gets the type of the event.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onFormulaChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onformulachanged-member)|Occurs when one or more formulas are changed in this worksheet.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onFormulaChanged](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onformulachanged-member)|Occurs when one or more formulas are changed in any worksheet of this collection.|
|[WorksheetFormulaChangedEventArgs](/javascript/api/excel/excel.worksheetformulachangedeventargs)|[formulaDetails](/javascript/api/excel/excel.worksheetformulachangedeventargs#excel-excel-worksheetformulachangedeventargs-formuladetails-member)|Gets an array of `FormulaChangedEventDetail` objects, which contain the details about the all of the changed formulas.|
||[source](/javascript/api/excel/excel.worksheetformulachangedeventargs#excel-excel-worksheetformulachangedeventargs-source-member)|The source of the event.|
||[type](/javascript/api/excel/excel.worksheetformulachangedeventargs#excel-excel-worksheetformulachangedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetformulachangedeventargs#excel-excel-worksheetformulachangedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the formula changed.|

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel?view=excel-js-1.13&preserve-view=true)
- [Excel JavaScript API requirement sets](excel-api-requirement-sets.md)
