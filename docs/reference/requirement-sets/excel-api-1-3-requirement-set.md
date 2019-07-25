---
title: Excel JavaScript API requirement set 1.3
description: 'Details about the ExcelApi 1.3 requirement set'
ms.date: 07/25/2019
ms.prod: excel
localization_priority: Normal
---

# What's new in Excel JavaScript API 1.3

ExcelApi 1.3 added support for data binding and basic PivotTable access.

## API list

To see a complete list of all APIs supported by this requirement set (including previously released APIs), [click here to see a version-specific of the API reference documentation](/javascript/api/excel?view=excel-js-1.3).

| Class | Fields | Description |
|:---|:---|:---|
|[Binding](/javascript/api/excel/excel.binding)|[delete()](/javascript/api/excel/excel.binding#delete--)|Deletes the binding.|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[add(range: Range \| string, bindingType: "Range" \| "Table" \| "Text", id: string)](/javascript/api/excel/excel.bindingcollection#add-range--bindingtype--id-)|Add a new binding to a particular Range.|
||[add(range: Range \| string, bindingType: Excel.BindingType, id: string)](/javascript/api/excel/excel.bindingcollection#add-range--bindingtype--id-)|Add a new binding to a particular Range.|
||[addFromNamedItem(name: string, bindingType: "Range" \| "Table" \| "Text", id: string)](/javascript/api/excel/excel.bindingcollection#addfromnameditem-name--bindingtype--id-)|Add a new binding based on a named item in the workbook.|
||[addFromNamedItem(name: string, bindingType: Excel.BindingType, id: string)](/javascript/api/excel/excel.bindingcollection#addfromnameditem-name--bindingtype--id-)|Add a new binding based on a named item in the workbook.|
||[addFromSelection(bindingType: "Range" \| "Table" \| "Text", id: string)](/javascript/api/excel/excel.bindingcollection#addfromselection-bindingtype--id-)|Add a new binding based on the current selection.|
||[addFromSelection(bindingType: Excel.BindingType, id: string)](/javascript/api/excel/excel.bindingcollection#addfromselection-bindingtype--id-)|Add a new binding based on the current selection.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[name](/javascript/api/excel/excel.pivottable#name)|Name of the PivotTable.|
||[worksheet](/javascript/api/excel/excel.pivottable#worksheet)|The worksheet containing the current PivotTable.|
||[refresh()](/javascript/api/excel/excel.pivottable#refresh--)|Refreshes the PivotTable.|
||[set(properties: Excel.PivotTable)](/javascript/api/excel/excel.pivottable#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.PivotTableUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.pivottable#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[getItem(name: string)](/javascript/api/excel/excel.pivottablecollection#getitem-name-)|Gets a PivotTable by name.|
||[items](/javascript/api/excel/excel.pivottablecollection#items)|Gets the loaded child items in this collection.|
||[refreshAll()](/javascript/api/excel/excel.pivottablecollection#refreshall--)|Refreshes all the pivot tables in the collection.|
|[PivotTableCollectionLoadOptions](/javascript/api/excel/excel.pivottablecollectionloadoptions)|[$all](/javascript/api/excel/excel.pivottablecollectionloadoptions#$all)||
||[name](/javascript/api/excel/excel.pivottablecollectionloadoptions#name)|For EACH ITEM in the collection: Name of the PivotTable.|
||[worksheet](/javascript/api/excel/excel.pivottablecollectionloadoptions#worksheet)|For EACH ITEM in the collection: The worksheet containing the current PivotTable.|
|[PivotTableData](/javascript/api/excel/excel.pivottabledata)|[name](/javascript/api/excel/excel.pivottabledata#name)|Name of the PivotTable.|
|[PivotTableLoadOptions](/javascript/api/excel/excel.pivottableloadoptions)|[$all](/javascript/api/excel/excel.pivottableloadoptions#$all)||
||[name](/javascript/api/excel/excel.pivottableloadoptions#name)|Name of the PivotTable.|
||[worksheet](/javascript/api/excel/excel.pivottableloadoptions#worksheet)|The worksheet containing the current PivotTable.|
|[PivotTableUpdateData](/javascript/api/excel/excel.pivottableupdatedata)|[name](/javascript/api/excel/excel.pivottableupdatedata#name)|Name of the PivotTable.|
|[Range](/javascript/api/excel/excel.range)|[getVisibleView()](/javascript/api/excel/excel.range#getvisibleview--)|Represents the visible rows of the current range.|
|[RangeView](/javascript/api/excel/excel.rangeview)|[formulas](/javascript/api/excel/excel.rangeview#formulas)|Represents the formula in A1-style notation.|
||[formulasLocal](/javascript/api/excel/excel.rangeview#formulaslocal)|Represents the formula in A1-style notation, in the user's language and number-formatting locale.  For example, the English "=SUM(A1, 1.5)" formula would become "=SUMME(A1; 1,5)" in German.|
||[formulasR1C1](/javascript/api/excel/excel.rangeview#formulasr1c1)|Represents the formula in R1C1-style notation.|
||[getRange()](/javascript/api/excel/excel.rangeview#getrange--)|Gets the parent range associated with the current RangeView.|
||[numberFormat](/javascript/api/excel/excel.rangeview#numberformat)|Represents Excel's number format code for the given cell.|
||[cellAddresses](/javascript/api/excel/excel.rangeview#celladdresses)|Represents the cell addresses of the RangeView. Read-only.|
||[columnCount](/javascript/api/excel/excel.rangeview#columncount)|Returns the number of visible columns. Read-only.|
||[index](/javascript/api/excel/excel.rangeview#index)|Returns a value that represents the index of the RangeView. Read-only.|
||[rowCount](/javascript/api/excel/excel.rangeview#rowcount)|Returns the number of visible rows. Read-only.|
||[rows](/javascript/api/excel/excel.rangeview#rows)|Represents a collection of range views associated with the range. Read-only.|
||[text](/javascript/api/excel/excel.rangeview#text)|Text values of the specified range. The Text value will not depend on the cell width. The # sign substitution that happens in Excel UI will not affect the text value returned by the API. Read-only.|
||[valueTypes](/javascript/api/excel/excel.rangeview#valuetypes)|Represents the type of data of each cell. Read-only.|
||[set(properties: Excel.RangeView)](/javascript/api/excel/excel.rangeview#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.RangeViewUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.rangeview#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
||[values](/javascript/api/excel/excel.rangeview#values)|Represents the raw values of the specified range view. The data returned could be of type string, number, or a boolean. Cells that contain an error will return the error string.|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[getItemAt(index: number)](/javascript/api/excel/excel.rangeviewcollection#getitemat-index-)|Gets a RangeView Row via its index. Zero-Indexed.|
||[items](/javascript/api/excel/excel.rangeviewcollection#items)|Gets the loaded child items in this collection.|
|[RangeViewCollectionLoadOptions](/javascript/api/excel/excel.rangeviewcollectionloadoptions)|[$all](/javascript/api/excel/excel.rangeviewcollectionloadoptions#$all)||
||[cellAddresses](/javascript/api/excel/excel.rangeviewcollectionloadoptions#celladdresses)|For EACH ITEM in the collection: Represents the cell addresses of the RangeView. Read-only.|
||[columnCount](/javascript/api/excel/excel.rangeviewcollectionloadoptions#columncount)|For EACH ITEM in the collection: Returns the number of visible columns. Read-only.|
||[formulas](/javascript/api/excel/excel.rangeviewcollectionloadoptions#formulas)|For EACH ITEM in the collection: Represents the formula in A1-style notation.|
||[formulasLocal](/javascript/api/excel/excel.rangeviewcollectionloadoptions#formulaslocal)|For EACH ITEM in the collection: Represents the formula in A1-style notation, in the user's language and number-formatting locale.  For example, the English "=SUM(A1, 1.5)" formula would become "=SUMME(A1; 1,5)" in German.|
||[formulasR1C1](/javascript/api/excel/excel.rangeviewcollectionloadoptions#formulasr1c1)|For EACH ITEM in the collection: Represents the formula in R1C1-style notation.|
||[index](/javascript/api/excel/excel.rangeviewcollectionloadoptions#index)|For EACH ITEM in the collection: Returns a value that represents the index of the RangeView. Read-only.|
||[numberFormat](/javascript/api/excel/excel.rangeviewcollectionloadoptions#numberformat)|For EACH ITEM in the collection: Represents Excel's number format code for the given cell.|
||[rowCount](/javascript/api/excel/excel.rangeviewcollectionloadoptions#rowcount)|For EACH ITEM in the collection: Returns the number of visible rows. Read-only.|
||[text](/javascript/api/excel/excel.rangeviewcollectionloadoptions#text)|For EACH ITEM in the collection: Text values of the specified range. The Text value will not depend on the cell width. The # sign substitution that happens in Excel UI will not affect the text value returned by the API. Read-only.|
||[valueTypes](/javascript/api/excel/excel.rangeviewcollectionloadoptions#valuetypes)|For EACH ITEM in the collection: Represents the type of data of each cell. Read-only.|
||[values](/javascript/api/excel/excel.rangeviewcollectionloadoptions#values)|For EACH ITEM in the collection: Represents the raw values of the specified range view. The data returned could be of type string, number, or a boolean. Cells that contain an error will return the error string.|
|[RangeViewData](/javascript/api/excel/excel.rangeviewdata)|[cellAddresses](/javascript/api/excel/excel.rangeviewdata#celladdresses)|Represents the cell addresses of the RangeView. Read-only.|
||[columnCount](/javascript/api/excel/excel.rangeviewdata#columncount)|Returns the number of visible columns. Read-only.|
||[formulas](/javascript/api/excel/excel.rangeviewdata#formulas)|Represents the formula in A1-style notation.|
||[formulasLocal](/javascript/api/excel/excel.rangeviewdata#formulaslocal)|Represents the formula in A1-style notation, in the user's language and number-formatting locale.  For example, the English "=SUM(A1, 1.5)" formula would become "=SUMME(A1; 1,5)" in German.|
||[formulasR1C1](/javascript/api/excel/excel.rangeviewdata#formulasr1c1)|Represents the formula in R1C1-style notation.|
||[index](/javascript/api/excel/excel.rangeviewdata#index)|Returns a value that represents the index of the RangeView. Read-only.|
||[numberFormat](/javascript/api/excel/excel.rangeviewdata#numberformat)|Represents Excel's number format code for the given cell.|
||[rowCount](/javascript/api/excel/excel.rangeviewdata#rowcount)|Returns the number of visible rows. Read-only.|
||[rows](/javascript/api/excel/excel.rangeviewdata#rows)|Represents a collection of range views associated with the range. Read-only.|
||[text](/javascript/api/excel/excel.rangeviewdata#text)|Text values of the specified range. The Text value will not depend on the cell width. The # sign substitution that happens in Excel UI will not affect the text value returned by the API. Read-only.|
||[valueTypes](/javascript/api/excel/excel.rangeviewdata#valuetypes)|Represents the type of data of each cell. Read-only.|
||[values](/javascript/api/excel/excel.rangeviewdata#values)|Represents the raw values of the specified range view. The data returned could be of type string, number, or a boolean. Cells that contain an error will return the error string.|
|[RangeViewLoadOptions](/javascript/api/excel/excel.rangeviewloadoptions)|[$all](/javascript/api/excel/excel.rangeviewloadoptions#$all)||
||[cellAddresses](/javascript/api/excel/excel.rangeviewloadoptions#celladdresses)|Represents the cell addresses of the RangeView. Read-only.|
||[columnCount](/javascript/api/excel/excel.rangeviewloadoptions#columncount)|Returns the number of visible columns. Read-only.|
||[formulas](/javascript/api/excel/excel.rangeviewloadoptions#formulas)|Represents the formula in A1-style notation.|
||[formulasLocal](/javascript/api/excel/excel.rangeviewloadoptions#formulaslocal)|Represents the formula in A1-style notation, in the user's language and number-formatting locale.  For example, the English "=SUM(A1, 1.5)" formula would become "=SUMME(A1; 1,5)" in German.|
||[formulasR1C1](/javascript/api/excel/excel.rangeviewloadoptions#formulasr1c1)|Represents the formula in R1C1-style notation.|
||[index](/javascript/api/excel/excel.rangeviewloadoptions#index)|Returns a value that represents the index of the RangeView. Read-only.|
||[numberFormat](/javascript/api/excel/excel.rangeviewloadoptions#numberformat)|Represents Excel's number format code for the given cell.|
||[rowCount](/javascript/api/excel/excel.rangeviewloadoptions#rowcount)|Returns the number of visible rows. Read-only.|
||[text](/javascript/api/excel/excel.rangeviewloadoptions#text)|Text values of the specified range. The Text value will not depend on the cell width. The # sign substitution that happens in Excel UI will not affect the text value returned by the API. Read-only.|
||[valueTypes](/javascript/api/excel/excel.rangeviewloadoptions#valuetypes)|Represents the type of data of each cell. Read-only.|
||[values](/javascript/api/excel/excel.rangeviewloadoptions#values)|Represents the raw values of the specified range view. The data returned could be of type string, number, or a boolean. Cells that contain an error will return the error string.|
|[RangeViewUpdateData](/javascript/api/excel/excel.rangeviewupdatedata)|[formulas](/javascript/api/excel/excel.rangeviewupdatedata#formulas)|Represents the formula in A1-style notation.|
||[formulasLocal](/javascript/api/excel/excel.rangeviewupdatedata#formulaslocal)|Represents the formula in A1-style notation, in the user's language and number-formatting locale.  For example, the English "=SUM(A1, 1.5)" formula would become "=SUMME(A1; 1,5)" in German.|
||[formulasR1C1](/javascript/api/excel/excel.rangeviewupdatedata#formulasr1c1)|Represents the formula in R1C1-style notation.|
||[numberFormat](/javascript/api/excel/excel.rangeviewupdatedata#numberformat)|Represents Excel's number format code for the given cell.|
||[values](/javascript/api/excel/excel.rangeviewupdatedata#values)|Represents the raw values of the specified range view. The data returned could be of type string, number, or a boolean. Cells that contain an error will return the error string.|
|[Table](/javascript/api/excel/excel.table)|[highlightFirstColumn](/javascript/api/excel/excel.table#highlightfirstcolumn)|Indicates whether the first column contains special formatting.|
||[highlightLastColumn](/javascript/api/excel/excel.table#highlightlastcolumn)|Indicates whether the last column contains special formatting.|
||[showBandedColumns](/javascript/api/excel/excel.table#showbandedcolumns)|Indicates whether the columns show banded formatting in which odd columns are highlighted differently from even ones to make reading the table easier.|
||[showBandedRows](/javascript/api/excel/excel.table#showbandedrows)|Indicates whether the rows show banded formatting in which odd rows are highlighted differently from even ones to make reading the table easier.|
||[showFilterButton](/javascript/api/excel/excel.table#showfilterbutton)|Indicates whether the filter buttons are visible at the top of each column header. Setting this is only allowed if the table contains a header row.|
|[TableCollectionLoadOptions](/javascript/api/excel/excel.tablecollectionloadoptions)|[highlightFirstColumn](/javascript/api/excel/excel.tablecollectionloadoptions#highlightfirstcolumn)|For EACH ITEM in the collection: Indicates whether the first column contains special formatting.|
||[highlightLastColumn](/javascript/api/excel/excel.tablecollectionloadoptions#highlightlastcolumn)|For EACH ITEM in the collection: Indicates whether the last column contains special formatting.|
||[showBandedColumns](/javascript/api/excel/excel.tablecollectionloadoptions#showbandedcolumns)|For EACH ITEM in the collection: Indicates whether the columns show banded formatting in which odd columns are highlighted differently from even ones to make reading the table easier.|
||[showBandedRows](/javascript/api/excel/excel.tablecollectionloadoptions#showbandedrows)|For EACH ITEM in the collection: Indicates whether the rows show banded formatting in which odd rows are highlighted differently from even ones to make reading the table easier.|
||[showFilterButton](/javascript/api/excel/excel.tablecollectionloadoptions#showfilterbutton)|For EACH ITEM in the collection: Indicates whether the filter buttons are visible at the top of each column header. Setting this is only allowed if the table contains a header row.|
|[TableData](/javascript/api/excel/excel.tabledata)|[highlightFirstColumn](/javascript/api/excel/excel.tabledata#highlightfirstcolumn)|Indicates whether the first column contains special formatting.|
||[highlightLastColumn](/javascript/api/excel/excel.tabledata#highlightlastcolumn)|Indicates whether the last column contains special formatting.|
||[showBandedColumns](/javascript/api/excel/excel.tabledata#showbandedcolumns)|Indicates whether the columns show banded formatting in which odd columns are highlighted differently from even ones to make reading the table easier.|
||[showBandedRows](/javascript/api/excel/excel.tabledata#showbandedrows)|Indicates whether the rows show banded formatting in which odd rows are highlighted differently from even ones to make reading the table easier.|
||[showFilterButton](/javascript/api/excel/excel.tabledata#showfilterbutton)|Indicates whether the filter buttons are visible at the top of each column header. Setting this is only allowed if the table contains a header row.|
|[TableLoadOptions](/javascript/api/excel/excel.tableloadoptions)|[highlightFirstColumn](/javascript/api/excel/excel.tableloadoptions#highlightfirstcolumn)|Indicates whether the first column contains special formatting.|
||[highlightLastColumn](/javascript/api/excel/excel.tableloadoptions#highlightlastcolumn)|Indicates whether the last column contains special formatting.|
||[showBandedColumns](/javascript/api/excel/excel.tableloadoptions#showbandedcolumns)|Indicates whether the columns show banded formatting in which odd columns are highlighted differently from even ones to make reading the table easier.|
||[showBandedRows](/javascript/api/excel/excel.tableloadoptions#showbandedrows)|Indicates whether the rows show banded formatting in which odd rows are highlighted differently from even ones to make reading the table easier.|
||[showFilterButton](/javascript/api/excel/excel.tableloadoptions#showfilterbutton)|Indicates whether the filter buttons are visible at the top of each column header. Setting this is only allowed if the table contains a header row.|
|[TableUpdateData](/javascript/api/excel/excel.tableupdatedata)|[highlightFirstColumn](/javascript/api/excel/excel.tableupdatedata#highlightfirstcolumn)|Indicates whether the first column contains special formatting.|
||[highlightLastColumn](/javascript/api/excel/excel.tableupdatedata#highlightlastcolumn)|Indicates whether the last column contains special formatting.|
||[showBandedColumns](/javascript/api/excel/excel.tableupdatedata#showbandedcolumns)|Indicates whether the columns show banded formatting in which odd columns are highlighted differently from even ones to make reading the table easier.|
||[showBandedRows](/javascript/api/excel/excel.tableupdatedata#showbandedrows)|Indicates whether the rows show banded formatting in which odd rows are highlighted differently from even ones to make reading the table easier.|
||[showFilterButton](/javascript/api/excel/excel.tableupdatedata#showfilterbutton)|Indicates whether the filter buttons are visible at the top of each column header. Setting this is only allowed if the table contains a header row.|
|[Workbook](/javascript/api/excel/excel.workbook)|[pivotTables](/javascript/api/excel/excel.workbook#pivottables)|Represents a collection of PivotTables associated with the workbook. Read-only.|
|[WorkbookData](/javascript/api/excel/excel.workbookdata)|[pivotTables](/javascript/api/excel/excel.workbookdata#pivottables)|Represents a collection of PivotTables associated with the workbook. Read-only.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[pivotTables](/javascript/api/excel/excel.worksheet#pivottables)|Collection of PivotTables that are part of the worksheet. Read-only.|
|[WorksheetData](/javascript/api/excel/excel.worksheetdata)|[pivotTables](/javascript/api/excel/excel.worksheetdata#pivottables)|Collection of PivotTables that are part of the worksheet. Read-only.|

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel&view=excel-js-1.3)
- [Excel JavaScript API requirement sets](./excel-api-requirement-sets.md)
