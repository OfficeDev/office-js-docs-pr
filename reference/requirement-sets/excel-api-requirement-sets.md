# Excel JavaScript API requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Specify Office hosts and API requirements](../../docs/overview/specify-office-hosts-and-api-requirements.md).

Excel add-ins run across multiple versions of Office, including Office 2016 for Windows, Office for iPad, Office for Mac, and Office Online. The following table lists the Excel requirement sets, the Office host applications that support that requirement set, and the build versions or number for those applications.

> For the requirement sets that are marked as *Beta*, use the specified (or later) version of the Office software and use the Beta library of the CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js. Entires not listed as *Beta* are generally available and you can continue to use Production CDN library: https://appsforoffice.microsoft.com/lib/1/hosted/office.js

|  Requirement set  |  Office 2016 for Windows*  |  Office 2016 for iPad  |  Office 2016 for Mac  | Office Online  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|
| ExcelApi 1.6 **Beta**  | Version 1702 (Build TBD) or later| Coming soon |  Coming soon| March 2016 | Coming soon|
| ExcelApi 1.5 **Beta**  | Version 1702 (Build TBD) or later| Coming soon |  Coming soon| March 2016 | Coming soon|
| ExcelApi 1.4 | Version 1701 (Build 7870.2024) or later| Coming soon |  Coming soon| March 2016 | Coming soon|
| ExcelApi 1.3  | Version 1608 (Build 7369.2055) or later| 1.27 or later |  15.27 or later| September 2016 | Version 1608 (Build 7601.6800) or later|
| ExcelApi 1.2  | Version 1601 (Build 6741.2088) or later | 1.21 or later | 15.22 or later| January 2016 ||
| ExcelApi 1.1  | Version 1509 (Build 4266.1001) or later | 1.19 or later | 15.20 or later| January 2016 ||

> **Note**: The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1 requirement set.

To find out more about versions, build numbers, and Office Online Server, see:

- [Version and build numbers of update channel releases for Office 365 clients](https://technet.microsoft.com/en-us/library/mt592918.aspx)
- [What version of Office am I using?](https://support.office.com/en-us/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19?ui=en-US&rs=en-US&ad=US&fromAR=1)
- [Where you can find the version and build number for an Office 365 client application](https://technet.microsoft.com/en-us/library/mt592918.aspx#Anchor_1)
- [Office Online Server overview](https://technet.microsoft.com/en-us/library/jj219437(v=office.16).aspx)

## Runtime requirement support check

During the runtime, add-ins can check if a particular host supports an API requirement set by doing the following-check: 

```js
if (Office.context.requirements.isSetSupported('ExcelApi', 1.3) === true) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

## Manifest based requirement support check

Use the Requirements element in the add-in manifest to specify critical requirement sets or API members that your add-in must use. If the Office host or platform doesn't support the requirement sets or API members specified in the Requirements element, the add-in won't run in that host or platform, and won't display in My Add-ins. Instead, we recommend that you make your add-in available on all platforms of an Office host, such as Excel for Windows, Excel Online, and Excel for iPad. To make your add-in available on all Office hosts and platforms, use runtime checks instead of the Requirements element.

The following code example shows an add-in that loads in all Office host applications that support ExcelApi requirement set, version 1.3.

```xml
<Requirements>
   <Sets DefaultMinVersion="1.3">
      <Set Name="ExcelApi" MinVersion="1.3"/>
   </Sets>
</Requirements>
```

## Office common API requirement sets
For information about common API requirement sets, see [Office common API requirement sets](office-add-in-requirement-sets.md).

## Upcoming Excel 1.6 Release Features

### Conditional formatting

Introduces [Conditional formating](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/conditionalformat.md) of a range. Allows follwoing types of conditional formatting:

* Color scale
* Data bar
* Icon set
* Custom

In addiiton:
* Returns the range the conditonal format is applied to.
* Removal of conditional formatting.
* Provides priority and stopifTrue capability
* Get collection of all conditional formatting on a given range.
* Clears all conditional formats active on the current specified range.

For API details, please refer to the Excel API [open specification](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec). 

## Upcoming Excel 1.5 Release Features

### Custom XML Part

* Addition of custom XML parts collection to workbook object.
* Get custom XML part using ID
* Get a new scoped collection of custom XML parts whose namespaces match the given namespace.
* Get XML string associated with a part.
* Provide id and namespace of a part.
* Adds a new custom XML part to the workbook.
* Set entire XML part.
* Delete a custom XML part.
* Delete an attribute with the given name from the element identified by xpath.
* Query the XML content by xpath.
* Insert, update and delete attribute.

**Reference implementation:** Please refer [here](https://github.com/mandren/Excel-CustomXMLPart-Demo) for a reference implementation that shows how custom XML parts can be used in an add-in.

### Others
* `range.getSurroundingRegion()` Returns a Range object that represents the surrounding region for this range. A surrounding region is a range bounded by any combination of blank rows and blank columns relative to this range.
* `getNextColumn()` and `getPreviousColumn()`, `getLast() on table column.
* `getActiveWorksheet()` on the workbook.
* `getRange(address: string)` off of workbook.
* `getBoundingRange(ranges: [])` Gets the smallest range object that encompasses the provided ranges. For example, the bounding range between "B2:C5" and "D10:E15" is "B2:E15".
* `getCount()` on various collections such as named item, worksheet, table, etc. to get number of items in a collection. `workbook.worksheets.getCount()`
* `getFirst()` and `getLast()` and get last on various collection such as tworksheet, able column, chart points, range view collection.
* `getNext()` and `getPrevious()` on worksheet, table column collection.
* `getRangeR1C1()` Gets the range object beginning at a particular row index and column index, and spanning a certain number of rows and columns.

For API details, please refer to the Excel API [open specification](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec). 

## What's new in Excel JavaScript API 1.4
The following are the new additions to the Excel JavaScript APIs in requirement set 1.3.

### Named item add and new properties

New properties
* `comment`
* `scope` worksheet or workbook scoped items
* `worksheet` returns the worksheet on which the named item is scoped to.

New Methods
* `add(name: string, reference: Range or string, comment: string)`Adds a new name to the collection of the given scope.
* `addFormulaLocal(name: string, formula: string, comment: string)` Adds a new name to the collection of the given scope using the user's locale for the formula.

### Settings API in in Excel namespace

[Setting](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_1.4_OpenSpec/reference/excel/setting.md) object represents a key-value pair of a setting persisted to the document. Now, we've added settings related APIs under Excel namespace. This doesn't offer net new functionality - however this make easy to remain in the promise based batched API syntax reduce the dependency on common API for Excel related tasks.

APIs include `getItem()` to get setting entry via the key, `add()` to add the specified key:value setting pair to the workbook.

### Others

* Set table column name (prior version only allows reading).
* Add table column to the end of the table (prior version only allows anywhere but last).
* Add multiple rows to a table at a time (prior version only allows 1 row at a time).
* `range.getColumnsAfter(count: number)` and `range.getColumnsBefore(count: number)` to get a certain number of columns to the right/left of the current Range object.
* Get item or null object function: This functionality allows getting object using a key. If the object does not exist, the returned object's isNullObject property will be true. This alows developers to check if an object exists or not without having to handle it thorugh exception handling. Available on worksheet, named-item, binding, chart series, etc.

`worksheet.GetItemOrNullObject()`

|Object| What is new| Description|Requirement set|
|:----|:----|:----|:----|
|[bindingCollection](../excel/bindingcollection.md)|_Method_ > [getCount()](../excel/bindingcollection.md#getcount)|Gets the number of bindings in the collection.|1.4|
|[bindingCollection](../excel/bindingcollection.md)|_Method_ > [getItemOrNullObject(id: string)](../excel/bindingcollection.md#getitemornullobjectid-string)|Gets a binding object by ID. If the binding object does not exist, will return a null object.|1.4|
|[chartCollection](../excel/chartcollection.md)|_Method_ > [getCount()](../excel/chartcollection.md#getcount)|Returns the number of charts in the worksheet.|1.4|
|[chartCollection](../excel/chartcollection.md)|_Method_ > [getItemOrNullObject(name: string)](../excel/chartcollection.md#getitemornullobjectname-string)|Gets a chart using its name. If there are multiple charts with the same name, the first one will be returned.|1.4|
|[chartPointsCollection](../excel/chartpointscollection.md)|_Method_ > [getCount()](../excel/chartpointscollection.md#getcount)|Returns the number of chart points in the series.|1.4|
|[chartSeriesCollection](../excel/chartseriescollection.md)|_Method_ > [getCount()](../excel/chartseriescollection.md#getcount)|Returns the number of series in the collection.|1.4|
|[namedItem](../excel/nameditem.md)|_Property_ > comment|Represents the comment associated with this name.|1.4|
|[namedItem](../excel/nameditem.md)|_Property_ > scope|Indicates whether the name is scoped to the workbook or to a specific worksheet. Read-only. Possible values are: Equal, Greater, GreaterEqual, Less, LessEqual, NotEqual.|1.4|
|[namedItem](../excel/nameditem.md)|_Relationship_ > worksheet|Returns the worksheet on which the named item is scoped to. Throws an error if the items is scoped to the workbook instead. Read-only.|1.4|
|[namedItem](../excel/nameditem.md)|_Relationship_ > worksheetOrNullObject|Returns the worksheet on which the named item is scoped to. Returns a null object if the item is scoped to the workbook instead. Read-only.|1.4|
|[namedItem](../excel/nameditem.md)|_Method_ > [delete()](../excel/nameditem.md#delete)|Deletes the given name.|1.4|
|[namedItem](../excel/nameditem.md)|_Method_ > [getRangeOrNullObject()](../excel/nameditem.md#getrangeornullobject)|Returns the range object that is associated with the name. Returns a null object if the named item's type is not a range.|1.4|
|[namedItemCollection](../excel/nameditemcollection.md)|_Method_ > [add(name: string, reference: Range or string, comment: string)](../excel/nameditemcollection.md#addname-string-reference-range-or-string-comment-string)|Adds a new name to the collection of the given scope.|1.4|
|[namedItemCollection](../excel/nameditemcollection.md)|_Method_ > [addFormulaLocal(name: string, formula: string, comment: string)](../excel/nameditemcollection.md#addformulalocalname-string-formula-string-comment-string)|Adds a new name to the collection of the given scope using the user's locale for the formula.|1.4|
|[namedItemCollection](../excel/nameditemcollection.md)|_Method_ > [getCount()](../excel/nameditemcollection.md#getcount)|Gets the number of named items in the collection.|1.4|
|[namedItemCollection](../excel/nameditemcollection.md)|_Method_ > [getItemOrNullObject(name: string)](../excel/nameditemcollection.md#getitemornullobjectname-string)|Gets a nameditem object using its name. If the nameditem object does not exist, will return a null object.|1.4|
|[pivotTableCollection](../excel/pivottablecollection.md)|_Method_ > [getCount()](../excel/pivottablecollection.md#getcount)|Gets the number of pivot tables in the collection.|1.4|
|[pivotTableCollection](../excel/pivottablecollection.md)|_Method_ > [getItemOrNullObject(name: string)](../excel/pivottablecollection.md#getitemornullobjectname-string)|Gets a PivotTable by name. If the PivotTable does not exist, will return a null object.|1.4|
|[range](../excel/range.md)|_Method_ > [getIntersectionOrNullObject(anotherRange: Range or string)](../excel/range.md#getintersectionornullobjectanotherrange-range-or-string)|Gets the range object that represents the rectangular intersection of the given ranges. If no intersection is found, will return a null object.|1.4|
|[range](../excel/range.md)|_Method_ > [getUsedRangeOrNullObject(valuesOnly: bool)](../excel/range.md#getusedrangeornullobjectvaluesonly-bool)|Returns the used range of the given range object. If there are no used cells within the range, this function will return a null object.|1.4|
|[rangeViewCollection](../excel/rangeviewcollection.md)|_Method_ > [getCount()](../excel/rangeviewcollection.md#getcount)|Gets the number of RangeView objects in the collection.|1.4|
|[setting](../excel/setting.md)|_Property_ > key|Returns the key that represents the id of the Setting. Read-only.|1.4|
|[setting](../excel/setting.md)|_Property_ > value|Represents the value stored for this setting.|1.4|
|[setting](../excel/setting.md)|_Method_ > [delete()](../excel/setting.md#delete)|Deletes the setting.|1.4|
|[settingCollection](../excel/settingcollection.md)|_Property_ > items|A collection of setting objects. Read-only.|1.4|
|[settingCollection](../excel/settingcollection.md)|_Method_ > [add(key: string, value: (any)[])](../excel/settingcollection.md#addkey-string-value-any)|Sets or adds the specified setting to the workbook.|1.4|
|[settingCollection](../excel/settingcollection.md)|_Method_ > [getCount()](../excel/settingcollection.md#getcount)|Gets the number of Settings in the collection.|1.4|
|[settingCollection](../excel/settingcollection.md)|_Method_ > [getItem(key: string)](../excel/settingcollection.md#getitemkey-string)|Gets a Setting entry via the key.|1.4|
|[settingCollection](../excel/settingcollection.md)|_Method_ > [getItemOrNullObject(key: string)](../excel/settingcollection.md#getitemornullobjectkey-string)|Gets a Setting entry via the key. If the Setting does not exist, will return a null object.|1.4|
|[settingsChangedEventArgs](../excel/settingschangedeventargs.md)|_Relationship_ > settings|Gets the Setting object that represents the binding that raised the SettingsChanged event|1.4|
|[tableCollection](../excel/tablecollection.md)|_Method_ > [getCount()](../excel/tablecollection.md#getcount)|Gets the number of tables in the collection.|1.4|
|[tableCollection](../excel/tablecollection.md)|_Method_ > [getItemOrNullObject(key: number or string)](../excel/tablecollection.md#getitemornullobjectkey-number-or-string)|Gets a table by Name or ID. If the table does not exist, will return a null object.|1.4|
|[tableColumnCollection](../excel/tablecolumncollection.md)|_Method_ > [getCount()](../excel/tablecolumncollection.md#getcount)|Gets the number of columns in the table.|1.4|
|[tableColumnCollection](../excel/tablecolumncollection.md)|_Method_ > [getItemOrNullObject(key: number or string)](../excel/tablecolumncollection.md#getitemornullobjectkey-number-or-string)|Gets a column object by Name or ID. If the column does not exist, will return a null object.|1.4|
|[tableRowCollection](../excel/tablerowcollection.md)|_Method_ > [getCount()](../excel/tablerowcollection.md#getcount)|Gets the number of rows in the table.|1.4|
|[workbook](../excel/workbook.md)|_Relationship_ > settings|Represents a collection of Settings associated with the workbook. Read-only.|1.4|
|[worksheet](../excel/worksheet.md)|_Relationship_ > names|Collection of names scoped to the current worksheet. Read-only.|1.4|
|[worksheet](../excel/worksheet.md)|_Method_ > [getUsedRangeOrNullObject(valuesOnly: bool)](../excel/worksheet.md#getusedrangeornullobjectvaluesonly-bool)|The used range is the smallest range that encompasses any cells that have a value or formatting assigned to them. If the entire worksheet is blank, this function will return a null object.|1.4|
|[worksheetCollection](../excel/worksheetcollection.md)|_Method_ > [getCount(visibleOnly: bool)](../excel/worksheetcollection.md#getcountvisibleonly-bool)|Gets the number of worksheets in the collection.|1.4|
|[worksheetCollection](../excel/worksheetcollection.md)|_Method_ > [getItemOrNullObject(key: string)](../excel/worksheetcollection.md#getitemornullobjectkey-string)|Gets a worksheet object using its Name or ID. If the worksheet does not exist, will return a null object.|1.4|



## What's new in Excel JavaScript API 1.3
The following are the new additions to the Excel JavaScript APIs in requirement set 1.3.

|Object| What's new| Description|Requirement set|
|:----|:----|:----|:----|
|[binding](../excel/binding.md)|_Method_ > [delete()](../excel/binding.md#delete)|Deletes the binding.|1.3|
|[bindingCollection](../excel/bindingcollection.md)|_Method_ > [add(range: Range or string, bindingType: string, id: string)](../excel/bindingcollection.md#addrange-range-or-string-bindingtype-string-id-string)|Add a new binding to a particular Range.|1.3|
|[bindingCollection](../excel/bindingcollection.md)|_Method_ > [addFromNamedItem(name: string, bindingType: string, id: string)](../excel/bindingcollection.md#addfromnameditemname-string-bindingtype-string-id-string)|Add a new binding based on a named item in the workbook.|1.3|
|[bindingCollection](../excel/bindingcollection.md)|_Method_ > [addFromSelection(bindingType: string, id: string)](../excel/bindingcollection.md#addfromselectionbindingtype-string-id-string)|Add a new binding based on the current selection.|1.3|
|[bindingCollection](../excel/bindingcollection.md)|_Method_ > [getItemOrNull(id: string)](../excel/bindingcollection.md#getitemornullid-string)|Gets a binding object by ID. If the binding object does not exist, the return object's isNull property will be true.|1.3|
|[chartCollection](../excel/chartcollection.md)|_Method_ > [getItemOrNull(name: string)](../excel/chartcollection.md#getitemornullname-string)|Gets a chart using its name. If there are multiple charts with the same name, the first one will be returned.|1.3|
|[namedItemCollection](../excel/nameditemcollection.md)|_Method_ > [getItemOrNull(name: string)](../excel/nameditemcollection.md#getitemornullname-string)|Gets a nameditem object using its name. If the nameditem object does not exist, the returned object's isNull property will be true.|1.3|
|[pivotTable](../excel/pivottable.md)|_Property_ > name|Name of the PivotTable.|1.3|
|[pivotTable](../excel/pivottable.md)|_Relationship_ > worksheet|The worksheet containing the current PivotTable. Read-only.|1.3|
|[pivotTable](../excel/pivottable.md)|_Method_ > [refresh()](../excel/pivottable.md#refresh)|Refreshes the PivotTable.|1.3|
|[pivotTableCollection](../excel/pivottablecollection.md)|_Property_ > items|A collection of pivotTable objects. Read-only.|1.3|
|[pivotTableCollection](../excel/pivottablecollection.md)|_Method_ > [getItem(name: string)](../excel/pivottablecollection.md#getitemname-string)|Gets a PivotTable by name.|1.3|
|[pivotTableCollection](../excel/pivottablecollection.md)|_Method_ > [getItemOrNull(name: string)](../excel/pivottablecollection.md#getitemornullname-string)|Gets a PivotTable by name. If the PivotTable does not exist, the return object's isNull property will be true.|1.3|
|[range](../excel/range.md)|_Method_ > [getIntersectionOrNull(anotherRange: Range or string)](../excel/range.md#getintersectionornullanotherrange-range-or-string)|Gets the range object that represents the rectangular intersection of the given ranges. If no intersection is found, will return a null object.|1.3|
|[range](../excel/range.md)|_Method_ > [getVisibleView()](../excel/range.md#getvisibleview)|Represents the visible rows of the current range.|1.3|
|[rangeView](../excel/rangeview.md)|_Property_ > cellAddresses|Represents the cell addresses of the RangeView. Read-only.|1.3|
|[rangeView](../excel/rangeview.md)|_Property_ > columnCount|Returns the number of visible columns. Read-only.|1.3|
|[rangeView](../excel/rangeview.md)|_Property_ > formulas|Represents the formula in A1-style notation.|1.3|
|[rangeView](../excel/rangeview.md)|_Property_ > formulasLocal|Represents the formula in A1-style notation, in the user's language and number-formatting locale.  For example, the English "=SUM(A1, introduced in 1.5)" formula would become "=SUMME(A1; 1,5)" in German.|1.3|
|[rangeView](../excel/rangeview.md)|_Property_ > formulasR1C1|Represents the formula in R1C1-style notation.|1.3|
|[rangeView](../excel/rangeview.md)|_Property_ > index|Returns a value that represents the index of the RangeView. Read-only.|1.3|
|[rangeView](../excel/rangeview.md)|_Property_ > numberFormat|Represents Excel's number format code for the given cell.|1.3|
|[rangeView](../excel/rangeview.md)|_Property_ > rowCount|Returns the number of visible rows. Read-only.|1.3|
|[rangeView](../excel/rangeview.md)|_Property_ > text|Text values of the specified range. The Text value will not depend on the cell width. The # sign substitution that happens in Excel UI will not affect the text value returned by the API. Read-only.|1.3|
|[rangeView](../excel/rangeview.md)|_Property_ > valueTypes|Represents the type of data of each cell. Read-only. Possible values are: Unknown, Empty, String, Integer, Double, Boolean, Error.|1.3|
|[rangeView](../excel/rangeview.md)|_Property_ > values|Represents the raw values of the specified range view. The data returned could be of type string, number, or a boolean. Cell that contain an error will return the error string.|1.3|
|[rangeView](../excel/rangeview.md)|_Relationship_ > rows|Represents a collection of range views associated with the range. Read-only.|1.3|
|[rangeView](../excel/rangeview.md)|_Method_ > [getRange()](../excel/rangeview.md#getrange)|Gets the parent range associated with the current RangeView.|1.3|
|[rangeViewCollection](../excel/rangeviewcollection.md)|_Property_ > items|A collection of rangeView objects. Read-only.|1.3|
|[rangeViewCollection](../excel/rangeviewcollection.md)|_Method_ > [getItemAt(index: number)](../excel/rangeviewcollection.md#getitematindex-number)|Gets a RangeView Row via it's index. Zero-Indexed.|1.3|
|[setting](../excel/setting.md)|_Property_ > key|Returns the key that represents the id of the Setting. Read-only.|1.3|
|[setting](../excel/setting.md)|_Method_ > [delete()](../excel/setting.md#delete)|Deletes the setting.|1.3|
|[settingCollection](../excel/settingcollection.md)|_Property_ > items|A collection of setting objects. Read-only.|1.3|
|[settingCollection](../excel/settingcollection.md)|_Method_ > [getItem(key: string)](../excel/settingcollection.md#getitemkey-string)|Gets a Setting entry via the key.|1.3|
|[settingCollection](../excel/settingcollection.md)|_Method_ > [getItemOrNull(key: string)](../excel/settingcollection.md#getitemornullkey-string)|Gets a Setting entry via the key. If the Setting does not exist, the returned object's isNull property will be true.|1.3|
|[settingCollection](../excel/settingcollection.md)|_Method_ > [set(key: string, value: string)](../excel/settingcollection.md#setkey-string-value-string)|Sets or adds the specified setting to the workbook.|1.3|
|[settingsChangedEventArgs](../excel/settingschangedeventargs.md)|_Relationship_ > settingCollection|Gets the Setting object that represents the binding that raised the SettingsChanged event|1.3|
|[table](../excel/table.md)|_Property_ > highlightFirstColumn|Indicates whether the first column contains special formatting.|1.3|
|[table](../excel/table.md)|_Property_ > highlightLastColumn|Indicates whether the last column contains special formatting.|1.3|
|[table](../excel/table.md)|_Property_ > showBandedColumns|Indicates whether the columns show banded formatting in which odd columns are highlighted differently from even ones to make reading the table easier.|1.3|
|[table](../excel/table.md)|_Property_ > showBandedRows|Indicates whether the rows show banded formatting in which odd rows are highlighted differently from even ones to make reading the table easier.|1.3|
|[table](../excel/table.md)|_Property_ > showFilterButton|Indicates whether the filter buttons are visible at the top of each column header. Setting this is only allowed if the table contains a header row.|1.3|
|[tableCollection](../excel/tablecollection.md)|_Method_ > [getItemOrNull(key: number or string)](../excel/tablecollection.md#getitemornullkey-number-or-string)|Gets a table by Name or ID. If the table does not exist, the return object's isNull property will be true.|1.3|
|[tableColumnCollection](../excel/tablecolumncollection.md)|_Method_ > [getItemOrNull(key: number or string)](../excel/tablecolumncollection.md#getitemornullkey-number-or-string)|Gets a column object by Name or ID. If the column does not exist, the returned object's isNull property will be true.|1.3|
|[workbook](../excel/workbook.md)|_Relationship_ > pivotTables|Represents a collection of PivotTables associated with the workbook. Read-only.|1.3|
|[workbook](../excel/workbook.md)|_Relationship_ > settings|Represents a collection of Settings associated with the workbook. Read-only.|1.3|
|[worksheet](../excel/worksheet.md)|_Relationship_ > pivotTables|Collection of PivotTables that are part of the worksheet. Read-only.|1.3|

## What's new in Excel JavaScript API 1.2
The following are the new additions to the Excel JavaScript APIs in requirement set 1.2.

|Object| What's new| Description|Requirement set|
|:----|:----|:----|:----|
|[chart](../excel/chart.md)|_Property_ > id|Gets a chart based on its position in the collection. Read-only.|1.2|
|[chart](../excel/chart.md)|_Relationship_ > worksheet|The worksheet containing the current chart. Read-only.|1.2|
|[chart](../excel/chart.md)|_Method_ > [getImage(height: number, width: number, fittingMode: string)](../excel/chart.md#getimageheight-number-width-number-fittingmode-string)|Renders the chart as a base64-encoded image by scaling the chart to fit the specified dimensions.|1.2|
|[filter](../excel/filter.md)|_Relationship_ > criteria|The currently applied filter on the given column. Read-only.|1.2|
|[filter](../excel/filter.md)|_Method_ > [apply(criteria: FilterCriteria)](../excel/filter.md#applycriteria-filtercriteria)|Apply the given filter criteria on the given column.|1.2|
|[filter](../excel/filter.md)|_Method_ > [applyBottomItemsFilter(count: number)](../excel/filter.md#applybottomitemsfiltercount-number)|Apply a "Bottom Item" filter to the column for the given number of elements.|1.2|
|[filter](../excel/filter.md)|_Method_ > [applyBottomPercentFilter(percent: number)](../excel/filter.md#applybottompercentfilterpercent-number)|Apply a "Bottom Percent" filter to the column for the given percentage of elements.|1.2|
|[filter](../excel/filter.md)|_Method_ > [applyCellColorFilter(color: string)](../excel/filter.md#applycellcolorfiltercolor-string)|Apply a "Cell Color" filter to the column for the given color.|1.2|
|[filter](../excel/filter.md)|_Method_ > [applyCustomFilter(criteria1: string, criteria2: string, oper: string)](../excel/filter.md#applycustomfiltercriteria1-string-criteria2-string-oper-string)|Apply a "Icon" filter to the column for the given criteria strings.|1.2|
|[filter](../excel/filter.md)|_Method_ > [applyDynamicFilter(criteria: string)](../excel/filter.md#applydynamicfiltercriteria-string)|Apply a "Dynamic" filter to the column.|1.2|
|[filter](../excel/filter.md)|_Method_ > [applyFontColorFilter(color: string)](../excel/filter.md#applyfontcolorfiltercolor-string)|Apply a "Font Color" filter to the column for the given color.|1.2|
|[filter](../excel/filter.md)|_Method_ > [applyIconFilter(icon: Icon)](../excel/filter.md#applyiconfiltericon-icon)|Apply a "Icon" filter to the column for the given icon.|1.2|
|[filter](../excel/filter.md)|_Method_ > [applyTopItemsFilter(count: number)](../excel/filter.md#applytopitemsfiltercount-number)|Apply a "Top Item" filter to the column for the given number of elements.|1.2|
|[filter](../excel/filter.md)|_Method_ > [applyTopPercentFilter(percent: number)](../excel/filter.md#applytoppercentfilterpercent-number)|Apply a "Top Percent" filter to the column for the given percentage of elements.|1.2|
|[filter](../excel/filter.md)|_Method_ > [applyValuesFilter(values: ()[])](../excel/filter.md#applyvaluesfiltervalues-)|Apply a "Values" filter to the column for the given values.|1.2|
|[filter](../excel/filter.md)|_Method_ > [clear()](../excel/filter.md#clear)|Clear the filter on the given column.|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_Property_ > color|The HTML color string used to filter cells. Used with "cellColor" and "fontColor" filtering.|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_Property_ > criterion1|The first criterion used to filter data. Used as an operator in the case of "custom" filtering.|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_Property_ > criterion2|The second criterion used to filter data. Only used as an operator in the case of "custom" filtering.|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_Property_ > dynamicCriteria|The dynamic criteria from the Excel.DynamicFilterCriteria set to apply on this column. Used with "dynamic" filtering. Possible values are: Unknown, AboveAverage, AllDatesInPeriodApril, AllDatesInPeriodAugust, AllDatesInPeriodDecember, AllDatesInPeriodFebruray, AllDatesInPeriodJanuary, AllDatesInPeriodJuly, AllDatesInPeriodJune, AllDatesInPeriodMarch, AllDatesInPeriodMay, AllDatesInPeriodNovember, AllDatesInPeriodOctober, AllDatesInPeriodQuarter1, AllDatesInPeriodQuarter2, AllDatesInPeriodQuarter3, AllDatesInPeriodQuarter4, AllDatesInPeriodSeptember, BelowAverage, LastMonth, LastQuarter, LastWeek, LastYear, NextMonth, NextQuarter, NextWeek, NextYear, ThisMonth, ThisQuarter, ThisWeek, ThisYear, Today, Tomorrow, YearToDate, Yesterday.|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_Property_ > filterOn|The property used by the filter to determine whether the values should stay visible. Possible values are: BottomItems, BottomPercent, CellColor, Dynamic, FontColor, Values, TopItems, TopPercent, Icon, Custom.|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_Property_ > operator|The operator used to combine criterion 1 and 2 when using "custom" filtering. Possible values are: And, Or.|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_Property_ > values|The set of values to be used as part of "values" filtering.|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_Relationship_ > icon|The icon used to filter cells. Used with "icon" filtering.|1.2|
|[filterDatetime](../excel/filterdatetime.md)|_Property_ > date|The date in ISO8601 format used to filter data.|1.2|
|[filterDatetime](../excel/filterdatetime.md)|_Property_ > specificity|How specific the date should be used to keep data. For example, if the date is 2005-04-02 and the specifity is set to "month", the filter operation will keep all rows with a date in the month of april 2009. Possible values are: Year, Monday, Day, Hour, Minute, Second.|1.2|
|[formatProtection](../excel/formatprotection.md)|_Property_ > formulaHidden|Indicates if Excel hides the formula for the cells in the range. A null value indicates that the entire range doesn't have uniform formula hidden setting.|1.2|
|[formatProtection](../excel/formatprotection.md)|_Property_ > locked|Indicates if Excel locks the cells in the object. A null value indicates that the entire range doesn't have uniform lock setting.|1.2|
|[icon](../excel/icon.md)|_Property_ > index|Represents the index of the icon in the given set.|1.2|
|[icon](../excel/icon.md)|_Property_ > set|Represents the set that the icon is part of. Possible values are: Invalid, ThreeArrows, ThreeArrowsGray, ThreeFlags, ThreeTrafficLights1, ThreeTrafficLights2, ThreeSigns, ThreeSymbols, ThreeSymbols2, FourArrows, FourArrowsGray, FourRedToBlack, FourRating, FourTrafficLights, FiveArrows, FiveArrowsGray, FiveRating, FiveQuarters, ThreeStars, ThreeTriangles, FiveBoxes.|1.2|
|[range](../excel/range.md)|_Property_ > columnHidden|Represents if all columns of the current range are hidden.|1.2|
|[range](../excel/range.md)|_Property_ > formulasR1C1|Represents the formula in R1C1-style notation.|1.2|
|[range](../excel/range.md)|_Property_ > hidden|Represents if all cells of the current range are hidden. Read-only.|1.2|
|[range](../excel/range.md)|_Property_ > rowHidden|Represents if all rows of the current range are hidden.|1.2|
|[range](../excel/range.md)|_Relationship_ > sort|Represents the range sort of the current range. Read-only.|1.2|
|[range](../excel/range.md)|_Method_ > [merge(across: bool)](../excel/range.md#mergeacross-bool)|Merge the range cells into one region in the worksheet.|1.2|
|[range](../excel/range.md)|_Method_ > [unmerge()](../excel/range.md#unmerge)|Unmerge the range cells into separate cells.|1.2|
|[rangeFormat](../excel/rangeformat.md)|_Property_ > columnWidth|Gets or sets the width of all colums within the range. If the column widths are not uniform, null will be returned.|1.2|
|[rangeFormat](../excel/rangeformat.md)|_Property_ > rowHeight|Gets or sets the height of all rows in the range. If the row heights are not uniform null will be returned.|1.2|
|[rangeFormat](../excel/rangeformat.md)|_Relationship_ > protection|Returns the format protection object for a range. Read-only.|1.2|
|[rangeFormat](../excel/rangeformat.md)|_Method_ > [autofitColumns()](../excel/rangeformat.md#autofitcolumns)|Changes the width of the columns of the current range to achieve the best fit, based on the current data in the columns.|1.2|
|[rangeFormat](../excel/rangeformat.md)|_Method_ > [autofitRows()](../excel/rangeformat.md#autofitrows)|Changes the height of the rows of the current range to achieve the best fit, based on the current data in the columns.|1.2|
|[rangeReference](../excel/rangereference.md)|_Property_ > address|Represents the visible rows of the current range.|1.2|
|[rangeSort](../excel/rangesort.md)|_Method_ > [apply(fields: SortField[], matchCase: bool, hasHeaders: bool, orientation: string, method: string)](../excel/rangesort.md#applyfields-sortfield-matchcase-bool-hasheaders-bool-orientation-string-method-string)|Perform a sort operation.|1.2|
|[sortField](../excel/sortfield.md)|_Property_ > ascending|Represents whether the sorting is done in an ascending fashion.|1.2|
|[sortField](../excel/sortfield.md)|_Property_ > color|Represents the color that is the target of the condition if the sorting is on font or cell color.|1.2|
|[sortField](../excel/sortfield.md)|_Property_ > dataOption|Represents additional sorting options for this field. Possible values are: Normal, TextAsNumber.|1.2|
|[sortField](../excel/sortfield.md)|_Property_ > key|Represents the column (or row, depending on the sort orientation) that the condition is on. Represented as an offset from the first column (or row).|1.2|
|[sortField](../excel/sortfield.md)|_Property_ > sortOn|Represents the type of sorting of this condition. Possible values are: Value, CellColor, FontColor, Icon.|1.2|
|[sortField](../excel/sortfield.md)|_Relationship_ > icon|Represents the icon that is the target of the condition if the sorting is on the cell's icon.|1.2|
|[table](../excel/table.md)|_Relationship_ > sort|Represents the sorting for the table. Read-only.|1.2|
|[table](../excel/table.md)|_Relationship_ > worksheet|The worksheet containing the current table. Read-only.|1.2|
|[table](../excel/table.md)|_Method_ > [clearFilters()](../excel/table.md#clearfilters)|Clears all the filters currently applied on the table.|1.2|
|[table](../excel/table.md)|_Method_ > [convertToRange()](../excel/table.md#converttorange)|Converts the table into a normal range of cells. All data is preserved.|1.2|
|[table](../excel/table.md)|_Method_ > [reapplyFilters()](../excel/table.md#reapplyfilters)|Reapplies all the filters currently on the table.|1.2|
|[tableColumn](../excel/tablecolumn.md)|_Relationship_ > filter|Retrieve the filter applied to the column. Read-only.|1.2|
|[tableSort](../excel/tablesort.md)|_Property_ > matchCase|Represents whether the casing impacted the last sort of the table. Read-only.|1.2|
|[tableSort](../excel/tablesort.md)|_Property_ > method|Represents Chinese character ordering method last used to sort the table. Read-only. Possible values are: PinYin, StrokeCount.|1.2|
|[tableSort](../excel/tablesort.md)|_Relationship_ > fields|Represents the current conditions used to last sort the table. Read-only.|1.2|
|[tableSort](../excel/tablesort.md)|_Method_ > [apply(fields: SortField[], matchCase: bool, method: string)](../excel/tablesort.md#applyfields-sortfield-matchcase-bool-method-string)|Perform a sort operation.|1.2|
|[tableSort](../excel/tablesort.md)|_Method_ > [clear()](../excel/tablesort.md#clear)|Clears the sorting that is currently on the table. While this doesn't modify the table's ordering, it clears the state of the header buttons.|1.2|
|[tableSort](../excel/tablesort.md)|_Method_ > [reapply()](../excel/tablesort.md#reapply)|Reapplies the current sorting parameters to the table.|1.2|
|[workbook](../excel/workbook.md)|_Relationship_ > functions|Represents Excel application instance that contains this workbook. Read-only.|1.2|
|[worksheet](../excel/worksheet.md)|_Relationship_ > protection|Returns sheet protection object for a worksheet. Read-only.|1.2|
|[worksheetProtection](../excel/worksheetprotection.md)|_Property_ > protected|Indicates if the worksheet is protected. Read-Only. Read-only.|1.2|
|[worksheetProtection](../excel/worksheetprotection.md)|_Relationship_ > options|Sheet protection options. Read-only.|1.2|
|[worksheetProtection](../excel/worksheetprotection.md)|_Method_ > [protect(options: WorksheetProtectionOptions)](../excel/worksheetprotection.md#protectoptions-worksheetprotectionoptions)|Protects a worksheet. Fails if the worksheet has been protected.|1.2|
|[worksheetProtection](../excel/worksheetprotection.md)|_Method_ > [unprotect()](../excel/worksheetprotection.md#unprotect)|Unprotects a worksheet.|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_Property_ > allowAutoFilter|Represents the worksheet protection option of allowing using auto filter feature.|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_Property_ > allowDeleteColumns|Represents the worksheet protection option of allowing deleting columns.|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_Property_ > allowDeleteRows|Represents the worksheet protection option of allowing deleting rows.|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_Property_ > allowFormatCells|Represents the worksheet protection option of allowing formatting cells.|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_Property_ > allowFormatColumns|Represents the worksheet protection option of allowing formatting columns.|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_Property_ > allowFormatRows|Represents the worksheet protection option of allowing formatting rows.|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_Property_ > allowInsertColumns|Represents the worksheet protection option of allowing inserting columns.|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_Property_ > allowInsertHyperlinks|Represents the worksheet protection option of allowing inserting hyperlinks.|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_Property_ > allowInsertRows|Represents the worksheet protection option of allowing inserting rows.|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_Property_ > allowPivotTables|Represents the worksheet protection option of allowing using PivotTable feature.|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_Property_ > allowSort|Represents the worksheet protection option of allowing using sort feature.|1.2|

## Excel JavaScript API 1.1
Excel JavaScript API 1.1 is the first version of the API. For details about the API,  see the Excel JavaScript API reference topics.  

## Additional resources

- [Specify Office hosts and API requirements](../../docs/overview/specify-office-hosts-and-api-requirements.md)
- [Office Add-ins XML manifest](../../docs/overview/add-in-manifests.md)
