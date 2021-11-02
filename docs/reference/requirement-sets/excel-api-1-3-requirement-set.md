---
title: Excel JavaScript API requirement set 1.3
description: 'Details about the ExcelApi 1.3 requirement set.'
ms.date: 11/09/2020
ms.prod: excel
ms.localizationpriority: medium
---

# What's new in Excel JavaScript API 1.3

ExcelApi 1.3 added support for data binding and basic PivotTable access.

## API list

The following table lists the APIs in Excel JavaScript API requirement set 1.3. To view API reference documentation for all APIs supported by Excel JavaScript API requirement set 1.3 or earlier, see [Excel APIs in requirement set 1.3 or earlier](/javascript/api/excel?view=excel-js-1.3&preserve-view=true).

| Class | Fields | Description |
|:---|:---|:---|
|[Binding](/javascript/api/excel/excel.binding)|[delete()](/javascript/api/excel/excel.binding#delete__)|Deletes the binding.|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[add(range: Range \| string, bindingType: Excel.BindingType, id: string)](/javascript/api/excel/excel.bindingcollection#add_range__bindingType__id_)|Add a new binding to a particular Range.|
||[addFromNamedItem(name: string, bindingType: Excel.BindingType, id: string)](/javascript/api/excel/excel.bindingcollection#addFromNamedItem_name__bindingType__id_)|Add a new binding based on a named item in the workbook.|
||[addFromSelection(bindingType: Excel.BindingType, id: string)](/javascript/api/excel/excel.bindingcollection#addFromSelection_bindingType__id_)|Add a new binding based on the current selection.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[name](/javascript/api/excel/excel.pivottable#name)|Name of the PivotTable.|
||[refresh()](/javascript/api/excel/excel.pivottable#refresh__)|Refreshes the PivotTable.|
||[worksheet](/javascript/api/excel/excel.pivottable#worksheet)|The worksheet containing the current PivotTable.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[getItem(name: string)](/javascript/api/excel/excel.pivottablecollection#getItem_name_)|Gets a PivotTable by name.|
||[items](/javascript/api/excel/excel.pivottablecollection#items)|Gets the loaded child items in this collection.|
||[refreshAll()](/javascript/api/excel/excel.pivottablecollection#refreshAll__)|Refreshes all the pivot tables in the collection.|
|[Range](/javascript/api/excel/excel.range)|[getVisibleView()](/javascript/api/excel/excel.range#getVisibleView__)|Represents the visible rows of the current range.|
|[RangeView](/javascript/api/excel/excel.rangeview)|[cellAddresses](/javascript/api/excel/excel.rangeview#cellAddresses)|Represents the cell addresses of the `RangeView`.|
||[columnCount](/javascript/api/excel/excel.rangeview#columnCount)|The number of visible columns.|
||[formulas](/javascript/api/excel/excel.rangeview#formulas)|Represents the formula in A1-style notation.|
||[formulasLocal](/javascript/api/excel/excel.rangeview#formulasLocal)|Represents the formula in A1-style notation, in the user's language and number-formatting locale.|
||[formulasR1C1](/javascript/api/excel/excel.rangeview#formulasR1C1)|Represents the formula in R1C1-style notation.|
||[getRange()](/javascript/api/excel/excel.rangeview#getRange__)|Gets the parent range associated with the current `RangeView`.|
||[index](/javascript/api/excel/excel.rangeview#index)|Returns a value that represents the index of the `RangeView`.|
||[numberFormat](/javascript/api/excel/excel.rangeview#numberFormat)|Represents Excel's number format code for the given cell.|
||[rowCount](/javascript/api/excel/excel.rangeview#rowCount)|The number of visible rows.|
||[rows](/javascript/api/excel/excel.rangeview#rows)|Represents a collection of range views associated with the range.|
||[text](/javascript/api/excel/excel.rangeview#text)|Text values of the specified range.|
||[valueTypes](/javascript/api/excel/excel.rangeview#valueTypes)|Represents the type of data of each cell.|
||[values](/javascript/api/excel/excel.rangeview#values)|Represents the raw values of the specified range view.|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[getItemAt(index: number)](/javascript/api/excel/excel.rangeviewcollection#getItemAt_index_)|Gets a `RangeView` row via its index.|
||[items](/javascript/api/excel/excel.rangeviewcollection#items)|Gets the loaded child items in this collection.|
|[Table](/javascript/api/excel/excel.table)|[highlightFirstColumn](/javascript/api/excel/excel.table#highlightFirstColumn)|Specifies if the first column contains special formatting.|
||[highlightLastColumn](/javascript/api/excel/excel.table#highlightLastColumn)|Specifies if the last column contains special formatting.|
||[showBandedColumns](/javascript/api/excel/excel.table#showBandedColumns)|Specifies if the columns show banded formatting in which odd columns are highlighted differently from even ones, to make reading the table easier.|
||[showBandedRows](/javascript/api/excel/excel.table#showBandedRows)|Specifies if the rows show banded formatting in which odd rows are highlighted differently from even ones, to make reading the table easier.|
||[showFilterButton](/javascript/api/excel/excel.table#showFilterButton)|Specifies if the filter buttons are visible at the top of each column header.|
|[Workbook](/javascript/api/excel/excel.workbook)|[pivotTables](/javascript/api/excel/excel.workbook#pivotTables)|Represents a collection of PivotTables associated with the workbook.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[pivotTables](/javascript/api/excel/excel.worksheet#pivotTables)|Collection of PivotTables that are part of the worksheet.|

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel?view=excel-js-1.3&preserve-view=true)
- [Excel JavaScript API requirement sets](excel-api-requirement-sets.md)
