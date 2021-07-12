---
title: Excel JavaScript API requirement set 1.13
description: 'Details about the ExcelApi 1.13 requirement set.'
ms.date: 07/09/2021
ms.prod: excel
localization_priority: Normal
---

# What's new in Excel JavaScript API 1.13

The ExcelApi 1.13 added a method to insert worksheets into a workbook from a Base64-encoded string and an event to detect workbook activation. It also increased support for formulas in ranges by adding APIs to track changes to formulas and locate a formula's direct dependent cells. Additionally, it expanded PivotTable support by adding PivotLayout APIs for alt text, style, and empty cell management.

| Feature area | Description | Relevant objects |
|:--- |:--- |:--- |
| [Formula changed events](../../excel/excel-add-ins-worksheets.md#detect-formula-changes) | Track changes to formulas, including the source and type of event that caused a change. | [Worksheet.onFormulaChanged](/javascript/api/excel/excel.worksheet#onFormulaChanged)|
| [Formula dependents](../../excel/excel-add-ins-ranges-precedents-dependents.md#get-the-direct-dependents-of-a-formula) | Locate the direct dependent cells of a formula. | [Range.getDirectDependents](/javascript/api/excel/excel.range#getDirectDependents__) |
| [Insert worksheets](../../excel//excel-add-ins-workbooks.md#insert-a-copy-of-an-existing-workbook-into-the-current-one) | Insert worksheets from another workbook into the current workbook as a Base64-encoded string. | [Workbook.insertWorksheetsFromBase64](/javascript/api/excel/excel.workbook#insertWorksheetsFromBase64_base64File__options_) |
| [PivotTable PivotLayout](../../excel/excel-add-ins-pivottables.md#other-pivotlayout-functions) | An expansion of the PivotLayout class, including new support for alt text and empty cell management. | [PivotLayout](/javascript/api/excel/excel.pivotlayout) |

## API list

The following table lists the APIs in Excel JavaScript API requirement set 1.13. To view API reference documentation for all APIs supported by Excel JavaScript API requirement set 1.13 or earlier, see [Excel APIs in requirement set 1.13 or earlier](/javascript/api/excel?view=excel-js-1.13&preserve-view=true).

| Class | Fields | Description |
|:---|:---|:---|
|[FormulaChangedEventDetail](/javascript/api/excel/excel.formulachangedeventdetail)|[cellAddress](/javascript/api/excel/excel.formulachangedeventdetail#celladdress)|The address of the cell that contains the changed formula.|
||[previousFormula](/javascript/api/excel/excel.formulachangedeventdetail#previousformula)|Represents the previous formula, before it was changed.|
|[InsertWorksheetOptions](/javascript/api/excel/excel.insertworksheetoptions)|[positionType](/javascript/api/excel/excel.insertworksheetoptions#positiontype)|The insert position, in the current workbook, of the new worksheets.|
||[relativeTo](/javascript/api/excel/excel.insertworksheetoptions#relativeto)|The worksheet in the current workbook that is referenced for the `WorksheetPositionType` parameter.|
||[sheetNamesToInsert](/javascript/api/excel/excel.insertworksheetoptions#sheetnamestoinsert)|The names of individual worksheets to insert.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[altTextDescription](/javascript/api/excel/excel.pivotlayout#alttextdescription)|The alt text description of the PivotTable.|
||[altTextTitle](/javascript/api/excel/excel.pivotlayout#alttexttitle)|The alt text title of the PivotTable.|
||[displayBlankLineAfterEachItem(display: boolean)](/javascript/api/excel/excel.pivotlayout#displayblanklineaftereachitem-display-)|Sets whether or not to display a blank line after each item.|
||[emptyCellText](/javascript/api/excel/excel.pivotlayout#emptycelltext)|The text that is automatically filled into any empty cell in the PivotTable if `fillEmptyCells == true`.|
||[fillEmptyCells](/javascript/api/excel/excel.pivotlayout#fillemptycells)|Specifies whether empty cells in the PivotTable should be populated with the `emptyCellText`.|
||[repeatAllItemLabels(repeatLabels: boolean)](/javascript/api/excel/excel.pivotlayout#repeatallitemlabels-repeatlabels-)|Sets the "repeat all item labels" setting across all fields in the PivotTable.|
||[showFieldHeaders](/javascript/api/excel/excel.pivotlayout#showfieldheaders)|Specifies whether the PivotTable displays field headers (field captions and filter drop-downs).|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[refreshOnOpen](/javascript/api/excel/excel.pivottable#refreshonopen)|Specifies whether the PivotTable refreshes when the workbook opens.|
|[Range](/javascript/api/excel/excel.range)|[getDirectDependents()](/javascript/api/excel/excel.range#getdirectdependents--)|Returns a `WorkbookRangeAreas` object that represents the range containing all the direct dependents of a cell in the same worksheet or in multiple worksheets.|
||[getExtendedRange(direction: Excel.KeyboardDirection, activeCell?: Range \| string)](/javascript/api/excel/excel.range#getextendedrange-direction--activecell-)|Returns a range object that includes the current range and up to the edge of the range, based on the provided direction.|
||[getMergedAreasOrNullObject()](/javascript/api/excel/excel.range#getmergedareasornullobject--)|Returns a RangeAreas object that represents the merged areas in this range.|
||[getRangeEdge(direction: Excel.KeyboardDirection, activeCell?: Range \| string)](/javascript/api/excel/excel.range#getrangeedge-direction--activecell-)|Returns a range object that is the edge cell of the data region that corresponds to the provided direction.|
|[Table](/javascript/api/excel/excel.table)|[resize(newRange: Range \| string)](/javascript/api/excel/excel.table#resize-newrange-)|Resize the table to the new range.|
|[Workbook](/javascript/api/excel/excel.workbook)|[insertWorksheetsFromBase64(base64File: string, options?: Excel.InsertWorksheetOptions)](/javascript/api/excel/excel.workbook#insertworksheetsfrombase64-base64file--options-)|Inserts the specified worksheets from a source workbook into the current workbook.|
||[onActivated](/javascript/api/excel/excel.workbook#onactivated)|Occurs when the the workbook is activated.|
|[WorkbookActivatedEventArgs](/javascript/api/excel/excel.workbookactivatedeventargs)|[type](/javascript/api/excel/excel.workbookactivatedeventargs#type)|Gets the type of the event.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onFormulaChanged](/javascript/api/excel/excel.worksheet#onformulachanged)|Occurs when one or more formulas are changed in this worksheet.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onFormulaChanged](/javascript/api/excel/excel.worksheetcollection#onformulachanged)|Occurs when one or more formulas are changed in any worksheet of this collection.|
|[WorksheetFormulaChangedEventArgs](/javascript/api/excel/excel.worksheetformulachangedeventargs)|[formulaDetails](/javascript/api/excel/excel.worksheetformulachangedeventargs#formuladetails)|Gets an array of `FormulaChangedEventDetail` objects, which contain the details about the all of the changed formulas.|
||[source](/javascript/api/excel/excel.worksheetformulachangedeventargs#source)|The source of the event.|
||[type](/javascript/api/excel/excel.worksheetformulachangedeventargs#type)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetformulachangedeventargs#worksheetid)|Gets the ID of the worksheet in which the formula changed.|

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel?view=excel-js-1.13&preserve-view=true)
- [Excel JavaScript API requirement sets](excel-api-requirement-sets.md)
