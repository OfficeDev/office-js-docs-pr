# Excel add-in requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine if an Office host supports APIs needed by the add-in. For more information, see [Specify Office hosts and API requirements](../docs/overview/specify-office-hosts-and-api-requirements.md).

## Requirement sets, platform availability, and build versions
Excel Add-ins run across multiple versions of Office including Office for Windows Desktop, Office Online, Office for the Mac, and Office for the iPad.

The following table lists the names of requirement sets, the Office host applications that support that requirement set, and the build versions.

For information about requirement sets for Outlook, see [Understanding Outlook API requirement sets](./outlook/tutorial-api-requirement-sets.md). To get a broad view of where add-ins are supported by Office host, see the [Office Add-in host and platform availability](https://dev.office.com/add-in-availability) page.

|  Set version  |  Office Online  |  Office 2016 for Windows  |  Office 2016 for iPad  |  Office 2016 for Mac  |
|:-----|-----|:-----|:-----|:-----|
| ExcelApi 1.3  | 7403.1000 | 7403.1000| 7403.1000 | 7403.1000|
| ExcelApi 1.2  | 7403.1000 | 7403.1000| 7403.1000 | 7403.1000|
| ExcelAPI 1.1  | 7403.1000 | 7403.1000| 7403.1000 | 7403.1000|

## Methods that aren't part of a requirement set


The following methods in the JavaScript API for Office aren't part of a requirement set. If your add-in requires any of these methods, use the  **Methods** and **Method** elements in the add-in's manifest to declare that they are required, or perform the runtime check using an if statement. For more information, see [Specify Office hosts and API requirements](../docs/overview/specify-office-hosts-and-api-requirements.md).

|**Method name**|**Office host support**|
|:-----|:-----|
|Bindings.addFromPromptAsync|Access web apps, Excel, and Excel Online|
|Document.getFilePropertiesAsync|Excel, Excel Online, Word, and PowerPoint|
|Document.getProjectFieldAsync|Project Standard 2013 and Project Professional 2013|
|Document.getResourceFieldAsync|Project Standard 2013 and Project Professional 2013|
|Document.getSelectedResourceAsync|Project Standard 2013 and Project Professional 2013|
|Document.getSelectedTaskAsync|Project Standard 2013 and Project Professional 2013|
|Document.getSelectedViewAsync|PowerPoint and PowerPoint Online|
|Document.getTaskAsync|Project Standard 2013 and Project Professional 2013|
|Document.getTaskFieldAsync|Project Standard 2013 and Project Professional 2013|
|Document.goToByIdAsync|Excel, Excel Online, Word, and PowerPoint|
|Settings.addHandlerAsync|Access web apps, Excel, Excel Online, Word, and PowerPoint|
|Settings.refreshAsync|Access web apps, Excel, Excel Online, Word, PowerPoint, and PowerPoint Online|
|Settings.removeHandlerAsync|Access web apps, Excel, Excel Online, Word, and PowerPoint|
|TableBinding.clearFormatsAsync|Excel, Excel Online|
|TableBinding.setFormatsAsync|Excel, Excel Online|
|TableBinding.setTableOptionsAsync|Excel, Excel Online|

## What's new in Excel JavaScript API 1.3 
The following are the new additions to the Excel JavaScript APIs in requirement set 1.3. 

|Object| What is new| Description|Feedback|
|:----|:----|:----|:----|
|[bindingCollection](reference/excel/bindingcollection.md)|_Method_ > [add(range: Range or string, bindingType: string, id: string)](reference/excel/bindingcollection.md#addrange-range-or-string-bindingtype-string-id-string)|Add a new binding to a particular Range.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-bindingCollection-add)|
|[bindingCollection](reference/excel/bindingcollection.md)|_Method_ > [addFromNamedItem(name: string, bindingType: string, id: string)](reference/excel/bindingcollection.md#addfromnameditemname-string-bindingtype-string-id-string)|Add a new binding based on a named item in the workbook.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-bindingCollection-addFromNamedItem)|
|[bindingCollection](reference/excel/bindingcollection.md)|_Method_ > [addFromSelection(bindingType: string, id: string)](reference/excel/bindingcollection.md#addfromselectionbindingtype-string-id-string)|Add a new binding based on the current selection.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-bindingCollection-addFromSelection)|
|[bindingCollection](reference/excel/bindingcollection.md)|_Method_ > [getItemOrNull(id: string)](reference/excel/bindingcollection.md#getitemornullid-string)|Gets a binding object by ID. If the binding object does not exist, the return object's isNull property will be true.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-bindingCollection-getItemOrNull)|
|[chartCollection](reference/excel/chartcollection.md)|_Method_ > [getItemOrNull(name: string)](reference/excel/chartcollection.md#getitemornullname-string)|Gets a chart using its name. If there are multiple charts with the same name, the first one will be returned.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-chartCollection-getItemOrNull)|
|[namedItemCollection](reference/excel/nameditemcollection.md)|_Method_ > [getItemOrNull(name: string)](reference/excel/nameditemcollection.md#getitemornullname-string)|Gets a nameditem object using its name. If the nameditem object does not exist, the returned object's isNull property will be true.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-namedItemCollection-getItemOrNull)|
|[pivotTable](reference/excel/pivottable.md)|_Property_ > name|Name of the PivotTable.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=pivotTable-name)|
|[pivotTable](reference/excel/pivottable.md)|_Relationship_ > worksheet|The worksheet containing the current PivotTable. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=pivotTable-worksheet)|
|[pivotTable](reference/excel/pivottable.md)|_Method_ > [refresh()](reference/excel/pivottable.md#refresh)|Refreshes the PivotTable.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-pivotTable-refresh)|
|[pivotTableCollection](reference/excel/pivottablecollection.md)|_Property_ > items|A collection of pivotTable objects. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=pivotTableCollection-items)|
|[pivotTableCollection](reference/excel/pivottablecollection.md)|_Method_ > [getItem(name: string)](reference/excel/pivottablecollection.md#getitemname-string)|Gets a PivotTable by name.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-pivotTableCollection-getItem)|
|[pivotTableCollection](reference/excel/pivottablecollection.md)|_Method_ > [getItemOrNull(name: string)](reference/excel/pivottablecollection.md#getitemornullname-string)|Gets a PivotTable by name. If the PivotTable does not exist, the return object's isNull property will be true.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-pivotTableCollection-getItemOrNull)|
|[pivotTableCollection](reference/excel/pivottablecollection.md)|_Method_ > [refreshAll()](reference/excel/pivottablecollection.md#refreshall)|Refreshes all the PivotTables in the collection.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-pivotTableCollection-refreshAll)|
|[range](reference/excel/range.md)|_Method_ > [getIntersectionOrNull(anotherRange: Range or string)](reference/excel/range.md#getintersectionornullanotherrange-range-or-string)|Gets the range object that represents the rectangular intersection of the given ranges. If no intersection is found, will return a null object.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-range-getIntersectionOrNull)|
|[range](reference/excel/range.md)|_Method_ > [getVisibleView()](reference/excel/range.md#getvisibleview)|Represents the visible rows of the current range.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-range-getVisibleView)|
|[rangeView](reference/excel/rangeview.md)|_Property_ > columnCount|Returns the number of visible columns. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=rangeView-columnCount)|
|[rangeView](reference/excel/rangeview.md)|_Property_ > formulas|Represents the formula in A1-style notation.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=rangeView-formulas)|
|[rangeView](reference/excel/rangeview.md)|_Property_ > formulasLocal|Represents the formula in A1-style notation, in the user's language and number-formatting locale.  For example, the English "=SUM(A1, 1.5)" formula would become "=SUMME(A1; 1,5)" in German.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=rangeView-formulasLocal)|
|[rangeView](reference/excel/rangeview.md)|_Property_ > formulasR1C1|Represents the formula in R1C1-style notation.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=rangeView-formulasR1C1)|
|[rangeView](reference/excel/rangeview.md)|_Property_ > numberFormat|Represents Excel's number format code for the given cell. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=rangeView-numberFormat)|
|[rangeView](reference/excel/rangeview.md)|_Property_ > rowCount|Returns the number of visible rows. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=rangeView-rowCount)|
|[rangeView](reference/excel/rangeview.md)|_Property_ > text|Text values of the specified range. The Text value will not depend on the cell width. The # sign substitution that happens in Excel UI will not affect the text value returned by the API. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=rangeView-text)|
|[rangeView](reference/excel/rangeview.md)|_Property_ > valueTypes|Represents the type of data of each cell. Read-only. Possible values are: Unknown, Empty, String, Integer, Double, Boolean, Error.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=rangeView-valueTypes)|
|[rangeView](reference/excel/rangeview.md)|_Property_ > values|Represents the raw values of the specified range view. The data returned could be of type string, number, or a boolean. Cell that contain an error will return the error string.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=rangeView-values)|
|[rangeView](reference/excel/rangeview.md)|_Relationship_ > rows|Represents a collection of range views associated with the range. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=rangeView-rows)|
|[rangeView](reference/excel/rangeview.md)|_Method_ > [getRange()](reference/excel/rangeview.md#getrange)|Gets the parent range associated with the current RangeView.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-rangeView-getRange)|
|[rangeViewCollection](reference/excel/rangeviewcollection.md)|_Property_ > items|A collection of rangeView objects. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=rangeViewCollection-items)|
|[rangeViewCollection](reference/excel/rangeviewcollection.md)|_Method_ > [getItem(index: number)](reference/excel/rangeviewcollection.md#getitemindex-number)|Gets a RangeView Row via it's index. Zero-Indexed.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-rangeViewCollection-getItem)|
|[table](reference/excel/table.md)|_Property_ > highlightFirstColumn|Indicates whether the first column contains special formatting.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-highlightFirstColumn)|
|[table](reference/excel/table.md)|_Property_ > highlightLastColumn|Indicates whether the last column contains special formatting.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-highlightLastColumn)|
|[table](reference/excel/table.md)|_Property_ > showBandedColumns|Indicates whether the columns show banded formatting in which odd columns are highlighted differently from even ones to make reading the table easier.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-showBandedColumns)|
|[table](reference/excel/table.md)|_Property_ > showBandedRows|Indicates whether the rows show banded formatting in which odd rows are highlighted differently from even ones to make reading the table easier.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-showBandedRows)|
|[table](reference/excel/table.md)|_Property_ > showFilterButton|Indicates whether the filter buttons are visible at the top of each column header. Setting this is only allowed if the table contains a header row.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-showFilterButton)|
|[tableCollection](reference/excel/tablecollection.md)|_Method_ > [getItemOrNull(key: number or string)](reference/excel/tablecollection.md#getitemornullkey-number-or-string)|Gets a table by Name or ID. If the table does not exist, the return object's isNull property will be true.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-tableCollection-getItemOrNull)|
|[tableColumnCollection](reference/excel/tablecolumncollection.md)|_Method_ > [getItemOrNull(key: number or string)](reference/excel/tablecolumncollection.md#getitemornullkey-number-or-string)|Gets a column object by Name or ID. If the column does not exist, the returned object's isNull property will be true.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-tableColumnCollection-getItemOrNull)|
|[workbook](reference/excel/workbook.md)|_Relationship_ > pivotTables|Represents a collection of PivotTables associated with the workbook. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=workbook-pivotTables)|
|[worksheet](reference/excel/worksheet.md)|_Property_ > visibility|The Visibility of the worksheet. Possible values are: Visible, Hidden, VeryHidden.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=worksheet-visibility)|
|[worksheet](reference/excel/worksheet.md)|_Relationship_ > pivotTables|Collection of PivotTables that are part of the worksheet. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=worksheet-pivotTables)|
|[worksheetCollection](reference/excel/worksheetcollection.md)|_Method_ > [getItemOrNull(key: string)](reference/excel/worksheetcollection.md#getitemornullkey-string)|Gets a worksheet object using its Name or ID. If the 
## Additional resources



- [Specify Office hosts and API requirements](../docs/overview/specify-office-hosts-and-api-requirements.md)
- [JavaScript API for Office reference](http://dev.office.com/reference/add-ins/javascript-api-for-office)
- [JavaScript API for Office reference](http://dev.office.com/reference/add-ins/javascript-api-for-office)
