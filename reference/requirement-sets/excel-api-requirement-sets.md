# Excel JavaScript API requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Specify Office hosts and API requirements](../docs/overview/specify-office-hosts-and-api-requirements.md).

Excel add-ins run across multiple versions of Office, including Office 2016 for Windows, Office for iPad, Office for Mac, and Office Online. The following table lists the Excel requirement sets, the Office host applications that support that requirement set, and the build versions or number for those applications. 

|  Requirement set  |  Office 2016 for Windows  |  Office 2016 for iPad  |  Office 2016 for Mac  | Office Online  |
|:-----|-----|:-----|:-----|:-----|
| ExcelApi 1.3  | Version 1608 or later| 1.27 or later |  15.27 or later| September 2016 | 
| ExcelApi 1.2  | Version 1601 or later | 1.21 or later | 15.22 or later| January 2016 |
| ExcelApi 1.1  | Shipped with Office 2016 version 1509 (Build 16.0.4266.1001) or later | 1.19 or later | 15.20 or later| January 2016 |

> **Note**: The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1 requirement set.

To find out more about versions and build numbers, see:

- [Version and build numbers of update channel releases for Office 365 clients](https://technet.microsoft.com/en-us/library/mt592918.aspx)
- [What version of Office am I using?](https://support.office.com/en-us/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19?ui=en-US&rs=en-US&ad=US&fromAR=1)
- [Where you can find the version and build number for an Office 365 client application](https://technet.microsoft.com/en-us/library/mt592918.aspx#Anchor_1)

## Office common API requirement sets
For information about common API requirement sets, see [Office common API requirement sets](office-add-in-requirement-sets.md).

## What's new in Excel JavaScript API 1.3 
The following are the new additions to the Excel JavaScript APIs in requirement set 1.3. 

|Object| What's new| Description|Requirement set|
|:----|:----|:----|:----|
|[binding](resources/binding.md)|_Method_ > [delete()](resources/binding.md#delete)|Deletes the binding.|1.3|
|[bindingCollection](resources/bindingcollection.md)|_Method_ > [add(range: Range or string, bindingType: string, id: string)](resources/bindingcollection.md#addrange-range-or-string-bindingtype-string-id-string)|Add a new binding to a particular Range.|1.3|
|[bindingCollection](resources/bindingcollection.md)|_Method_ > [addFromNamedItem(name: string, bindingType: string, id: string)](resources/bindingcollection.md#addfromnameditemname-string-bindingtype-string-id-string)|Add a new binding based on a named item in the workbook.|1.3|
|[bindingCollection](resources/bindingcollection.md)|_Method_ > [addFromSelection(bindingType: string, id: string)](resources/bindingcollection.md#addfromselectionbindingtype-string-id-string)|Add a new binding based on the current selection.|1.3|
|[bindingCollection](resources/bindingcollection.md)|_Method_ > [getItemOrNull(id: string)](resources/bindingcollection.md#getitemornullid-string)|Gets a binding object by ID. If the binding object does not exist, the return object's isNull property will be true.|1.3|
|[chartCollection](resources/chartcollection.md)|_Method_ > [getItemOrNull(name: string)](resources/chartcollection.md#getitemornullname-string)|Gets a chart using its name. If there are multiple charts with the same name, the first one will be returned.|1.3|
|[namedItemCollection](resources/nameditemcollection.md)|_Method_ > [getItemOrNull(name: string)](resources/nameditemcollection.md#getitemornullname-string)|Gets a nameditem object using its name. If the nameditem object does not exist, the returned object's isNull property will be true.|1.3|
|[pivotTable](resources/pivottable.md)|_Property_ > name|Name of the PivotTable.|1.3|
|[pivotTable](resources/pivottable.md)|_Relationship_ > worksheet|The worksheet containing the current PivotTable. Read-only.|1.3|
|[pivotTable](resources/pivottable.md)|_Method_ > [refresh()](resources/pivottable.md#refresh)|Refreshes the PivotTable.|1.3|
|[pivotTableCollection](resources/pivottablecollection.md)|_Property_ > items|A collection of pivotTable objects. Read-only.|1.3|
|[pivotTableCollection](resources/pivottablecollection.md)|_Method_ > [getItem(name: string)](resources/pivottablecollection.md#getitemname-string)|Gets a PivotTable by name.|1.3|
|[pivotTableCollection](resources/pivottablecollection.md)|_Method_ > [getItemOrNull(name: string)](resources/pivottablecollection.md#getitemornullname-string)|Gets a PivotTable by name. If the PivotTable does not exist, the return object's isNull property will be true.|1.3|
|[range](resources/range.md)|_Method_ > [getIntersectionOrNull(anotherRange: Range or string)](resources/range.md#getintersectionornullanotherrange-range-or-string)|Gets the range object that represents the rectangular intersection of the given ranges. If no intersection is found, will return a null object.|1.3|
|[range](resources/range.md)|_Method_ > [getVisibleView()](resources/range.md#getvisibleview)|Represents the visible rows of the current range.|1.3|
|[rangeView](resources/rangeview.md)|_Property_ > cellAddresses|Represents the cell addresses of the RangeView. Read-only.|1.3|
|[rangeView](resources/rangeview.md)|_Property_ > columnCount|Returns the number of visible columns. Read-only.|1.3|
|[rangeView](resources/rangeview.md)|_Property_ > formulas|Represents the formula in A1-style notation.|1.3|
|[rangeView](resources/rangeview.md)|_Property_ > formulasLocal|Represents the formula in A1-style notation, in the user's language and number-formatting locale.  For example, the English "=SUM(A1, introduced in 1.5)" formula would become "=SUMME(A1; 1,5)" in German.|1.3|
|[rangeView](resources/rangeview.md)|_Property_ > formulasR1C1|Represents the formula in R1C1-style notation.|1.3|
|[rangeView](resources/rangeview.md)|_Property_ > index|Returns a value that represents the index of the RangeView. Read-only.|1.3|
|[rangeView](resources/rangeview.md)|_Property_ > numberFormat|Represents Excel's number format code for the given cell.|1.3|
|[rangeView](resources/rangeview.md)|_Property_ > rowCount|Returns the number of visible rows. Read-only.|1.3|
|[rangeView](resources/rangeview.md)|_Property_ > text|Text values of the specified range. The Text value will not depend on the cell width. The # sign substitution that happens in Excel UI will not affect the text value returned by the API. Read-only.|1.3|
|[rangeView](resources/rangeview.md)|_Property_ > valueTypes|Represents the type of data of each cell. Read-only. Possible values are: Unknown, Empty, String, Integer, Double, Boolean, Error.|1.3|
|[rangeView](resources/rangeview.md)|_Property_ > values|Represents the raw values of the specified range view. The data returned could be of type string, number, or a boolean. Cell that contain an error will return the error string.|1.3|
|[rangeView](resources/rangeview.md)|_Relationship_ > rows|Represents a collection of range views associated with the range. Read-only.|1.3|
|[rangeView](resources/rangeview.md)|_Method_ > [getRange()](resources/rangeview.md#getrange)|Gets the parent range associated with the current RangeView.|1.3|
|[rangeViewCollection](resources/rangeviewcollection.md)|_Property_ > items|A collection of rangeView objects. Read-only.|1.3|
|[rangeViewCollection](resources/rangeviewcollection.md)|_Method_ > [getItemAt(index: number)](resources/rangeviewcollection.md#getitematindex-number)|Gets a RangeView Row via it's index. Zero-Indexed.|1.3|
|[setting](resources/setting.md)|_Property_ > key|Returns the key that represents the id of the Setting. Read-only.|1.3|
|[setting](resources/setting.md)|_Method_ > [delete()](resources/setting.md#delete)|Deletes the setting.|1.3|
|[settingCollection](resources/settingcollection.md)|_Property_ > items|A collection of setting objects. Read-only.|1.3|
|[settingCollection](resources/settingcollection.md)|_Method_ > [getItem(key: string)](resources/settingcollection.md#getitemkey-string)|Gets a Setting entry via the key.|1.3|
|[settingCollection](resources/settingcollection.md)|_Method_ > [getItemOrNull(key: string)](resources/settingcollection.md#getitemornullkey-string)|Gets a Setting entry via the key. If the Setting does not exist, the returned object's isNull property will be true.|1.3|
|[settingCollection](resources/settingcollection.md)|_Method_ > [set(key: string, value: string)](resources/settingcollection.md#setkey-string-value-string)|Sets or adds the specified setting to the workbook.|1.3|
|[settingsChangedEventArgs](resources/settingschangedeventargs.md)|_Relationship_ > settingCollection|Gets the Setting object that represents the binding that raised the SettingsChanged event|1.3|
|[table](resources/table.md)|_Property_ > highlightFirstColumn|Indicates whether the first column contains special formatting.|1.3|
|[table](resources/table.md)|_Property_ > highlightLastColumn|Indicates whether the last column contains special formatting.|1.3|
|[table](resources/table.md)|_Property_ > showBandedColumns|Indicates whether the columns show banded formatting in which odd columns are highlighted differently from even ones to make reading the table easier.|1.3|
|[table](resources/table.md)|_Property_ > showBandedRows|Indicates whether the rows show banded formatting in which odd rows are highlighted differently from even ones to make reading the table easier.|1.3|
|[table](resources/table.md)|_Property_ > showFilterButton|Indicates whether the filter buttons are visible at the top of each column header. Setting this is only allowed if the table contains a header row.|1.3|
|[tableCollection](resources/tablecollection.md)|_Method_ > [getItemOrNull(key: number or string)](resources/tablecollection.md#getitemornullkey-number-or-string)|Gets a table by Name or ID. If the table does not exist, the return object's isNull property will be true.|1.3|
|[tableColumnCollection](resources/tablecolumncollection.md)|_Method_ > [getItemOrNull(key: number or string)](resources/tablecolumncollection.md#getitemornullkey-number-or-string)|Gets a column object by Name or ID. If the column does not exist, the returned object's isNull property will be true.|1.3|
|[workbook](resources/workbook.md)|_Relationship_ > pivotTables|Represents a collection of PivotTables associated with the workbook. Read-only.|1.3|
|[workbook](resources/workbook.md)|_Relationship_ > settings|Represents a collection of Settings associated with the workbook. Read-only.|1.3|
|[worksheet](resources/worksheet.md)|_Relationship_ > pivotTables|Collection of PivotTables that are part of the worksheet. Read-only.|1.3|

## What's new in Excel JavaScript API 1.2
The following are the new additions to the Excel JavaScript APIs in requirement set 1.2. 

|Object| What's new| Description|Requirement set|
|:----|:----|:----|:----|
|[bindingDataChangedEventArgs](resources/bindingdatachangedeventargs.md)|_Relationship_ > binding|Gets the Binding object that represents the binding that raised the DataChanged event.|1.2, introduced in 1.3|
|[bindingSelectionChangedEventArgs](resources/bindingselectionchangedeventargs.md)|_Property_ > columnCount|Gets the number of columns selected.|1.2, introduced in 1.3|
|[bindingSelectionChangedEventArgs](resources/bindingselectionchangedeventargs.md)|_Property_ > rowCount|Gets the number of rows selected.|1.2, introduced in 1.3|
|[bindingSelectionChangedEventArgs](resources/bindingselectionchangedeventargs.md)|_Property_ > startColumn|Gets the index of the first column of the selection (zero-based).|1.2, introduced in 1.3|
|[bindingSelectionChangedEventArgs](resources/bindingselectionchangedeventargs.md)|_Property_ > startRow|Gets the index of the first row of the selection (zero-based).|1.2, introduced in 1.3|
|[bindingSelectionChangedEventArgs](resources/bindingselectionchangedeventargs.md)|_Relationship_ > binding|Gets the Binding object that represents the binding that raised the SelectionChanged event.|1.2, introduced in 1.3|
|[chart](resources/chart.md)|_Property_ > id|Gets a chart based on its position in the collection. Read-only.|1.2|
|[chart](resources/chart.md)|_Relationship_ > worksheet|The worksheet containing the current chart. Read-only.|1.2|
|[chart](resources/chart.md)|_Method_ > [getImage(height: number, width: number, fittingMode: string)](resources/chart.md#getimageheight-number-width-number-fittingmode-string)|Renders the chart as a base64-encoded image by scaling the chart to fit the specified dimensions.|1.2|
|[filter](resources/filter.md)|_Relationship_ > criteria|The currently applied filter on the given column. Read-only.|1.2|
|[filter](resources/filter.md)|_Method_ > [apply(criteria: FilterCriteria)](resources/filter.md#applycriteria-filtercriteria)|Apply the given filter criteria on the given column.|1.2|
|[filter](resources/filter.md)|_Method_ > [applyBottomItemsFilter(count: number)](resources/filter.md#applybottomitemsfiltercount-number)|Apply a "Bottom Item" filter to the column for the given number of elements.|1.2|
|[filter](resources/filter.md)|_Method_ > [applyBottomPercentFilter(percent: number)](resources/filter.md#applybottompercentfilterpercent-number)|Apply a "Bottom Percent" filter to the column for the given percentage of elements.|1.2|
|[filter](resources/filter.md)|_Method_ > [applyCellColorFilter(color: string)](resources/filter.md#applycellcolorfiltercolor-string)|Apply a "Cell Color" filter to the column for the given color.|1.2|
|[filter](resources/filter.md)|_Method_ > [applyCustomFilter(criteria1: string, criteria2: string, oper: string)](resources/filter.md#applycustomfiltercriteria1-string-criteria2-string-oper-string)|Apply a "Icon" filter to the column for the given criteria strings.|1.2|
|[filter](resources/filter.md)|_Method_ > [applyDynamicFilter(criteria: string)](resources/filter.md#applydynamicfiltercriteria-string)|Apply a "Dynamic" filter to the column.|1.2|
|[filter](resources/filter.md)|_Method_ > [applyFontColorFilter(color: string)](resources/filter.md#applyfontcolorfiltercolor-string)|Apply a "Font Color" filter to the column for the given color.|1.2|
|[filter](resources/filter.md)|_Method_ > [applyIconFilter(icon: Icon)](resources/filter.md#applyiconfiltericon-icon)|Apply a "Icon" filter to the column for the given icon.|1.2|
|[filter](resources/filter.md)|_Method_ > [applyTopItemsFilter(count: number)](resources/filter.md#applytopitemsfiltercount-number)|Apply a "Top Item" filter to the column for the given number of elements.|1.2|
|[filter](resources/filter.md)|_Method_ > [applyTopPercentFilter(percent: number)](resources/filter.md#applytoppercentfilterpercent-number)|Apply a "Top Percent" filter to the column for the given percentage of elements.|1.2|
|[filter](resources/filter.md)|_Method_ > [applyValuesFilter(values: ()[])](resources/filter.md#applyvaluesfiltervalues-)|Apply a "Values" filter to the column for the given values.|1.2|
|[filter](resources/filter.md)|_Method_ > [clear()](resources/filter.md#clear)|Clear the filter on the given column.|1.2|
|[filterCriteria](resources/filtercriteria.md)|_Property_ > color|The HTML color string used to filter cells. Used with "cellColor" and "fontColor" filtering.|1.2|
|[filterCriteria](resources/filtercriteria.md)|_Property_ > criterion1|The first criterion used to filter data. Used as an operator in the case of "custom" filtering.|1.2|
|[filterCriteria](resources/filtercriteria.md)|_Property_ > criterion2|The second criterion used to filter data. Only used as an operator in the case of "custom" filtering.|1.2|
|[filterCriteria](resources/filtercriteria.md)|_Property_ > dynamicCriteria|The dynamic criteria from the Excel.DynamicFilterCriteria set to apply on this column. Used with "dynamic" filtering. Possible values are: Unknown, AboveAverage, AllDatesInPeriodApril, AllDatesInPeriodAugust, AllDatesInPeriodDecember, AllDatesInPeriodFebruray, AllDatesInPeriodJanuary, AllDatesInPeriodJuly, AllDatesInPeriodJune, AllDatesInPeriodMarch, AllDatesInPeriodMay, AllDatesInPeriodNovember, AllDatesInPeriodOctober, AllDatesInPeriodQuarter1, AllDatesInPeriodQuarter2, AllDatesInPeriodQuarter3, AllDatesInPeriodQuarter4, AllDatesInPeriodSeptember, BelowAverage, LastMonth, LastQuarter, LastWeek, LastYear, NextMonth, NextQuarter, NextWeek, NextYear, ThisMonth, ThisQuarter, ThisWeek, ThisYear, Today, Tomorrow, YearToDate, Yesterday.|1.2|
|[filterCriteria](resources/filtercriteria.md)|_Property_ > filterOn|The property used by the filter to determine whether the values should stay visible. Possible values are: BottomItems, BottomPercent, CellColor, Dynamic, FontColor, Values, TopItems, TopPercent, Icon, Custom.|1.2|
|[filterCriteria](resources/filtercriteria.md)|_Property_ > operator|The operator used to combine criterion 1 and 2 when using "custom" filtering. Possible values are: And, Or.|1.2|
|[filterCriteria](resources/filtercriteria.md)|_Property_ > values|The set of values to be used as part of "values" filtering.|1.2|
|[filterCriteria](resources/filtercriteria.md)|_Relationship_ > icon|The icon used to filter cells. Used with "icon" filtering.|1.2|
|[filterDatetime](resources/filterdatetime.md)|_Property_ > date|The date in ISO8601 format used to filter data.|1.2|
|[filterDatetime](resources/filterdatetime.md)|_Property_ > specificity|How specific the date should be used to keep data. For example, if the date is 2005-04-02 and the specifity is set to "month", the filter operation will keep all rows with a date in the month of april 2009. Possible values are: Year, Monday, Day, Hour, Minute, Second.|1.2|
|[formatProtection](resources/formatprotection.md)|_Property_ > formulaHidden|Indicates if Excel hides the formula for the cells in the range. A null value indicates that the entire range doesn't have uniform formula hidden setting.|1.2|
|[formatProtection](resources/formatprotection.md)|_Property_ > locked|Indicates if Excel locks the cells in the object. A null value indicates that the entire range doesn't have uniform lock setting.|1.2|
|[icon](resources/icon.md)|_Property_ > index|Represents the index of the icon in the given set.|1.2|
|[icon](resources/icon.md)|_Property_ > set|Represents the set that the icon is part of. Possible values are: Invalid, ThreeArrows, ThreeArrowsGray, ThreeFlags, ThreeTrafficLights1, ThreeTrafficLights2, ThreeSigns, ThreeSymbols, ThreeSymbols2, FourArrows, FourArrowsGray, FourRedToBlack, FourRating, FourTrafficLights, FiveArrows, FiveArrowsGray, FiveRating, FiveQuarters, ThreeStars, ThreeTriangles, FiveBoxes.|1.2|
|[range](resources/range.md)|_Property_ > columnHidden|Represents if all columns of the current range are hidden.|1.2|
|[range](resources/range.md)|_Property_ > formulasR1C1|Represents the formula in R1C1-style notation.|1.2|
|[range](resources/range.md)|_Property_ > hidden|Represents if all cells of the current range are hidden. Read-only.|1.2|
|[range](resources/range.md)|_Property_ > rowHidden|Represents if all rows of the current range are hidden.|1.2|
|[range](resources/range.md)|_Relationship_ > sort|Represents the range sort of the current range. Read-only.|1.2|
|[range](resources/range.md)|_Method_ > [getColumnsAfter(count: number)](resources/range.md#getcolumnsaftercount-number)|Gets a certain number of columns to the right of the current Range object.|1.2, introduced in 1.3|
|[range](resources/range.md)|_Method_ > [getColumnsBefore(count: number)](resources/range.md#getcolumnsbeforecount-number)|Gets a certain number of columns to the left of the current Range object.|1.2, introduced in 1.3|
|[range](resources/range.md)|_Method_ > [getResizedRange(deltaRows: number, deltaColumns: number)](resources/range.md#getresizedrangedeltarows-number-deltacolumns-number)|Gets a Range object similar to the current Range object, but with its bottom-right corner expanded (or contracted) by some number of rows and columns.|1.2, introduced in 1.3|
|[range](resources/range.md)|_Method_ > [getRowsAbove(count: number)](resources/range.md#getrowsabovecount-number)|Gets a certain number of rows above the current Range object.|1.2, introduced in 1.3|
|[range](resources/range.md)|_Method_ > [getRowsBelow(count: number)](resources/range.md#getrowsbelowcount-number)|Gets a certain number of rows below the current Range object.|1.2, introduced in 1.3|
|[range](resources/range.md)|_Method_ > [merge(across: bool)](resources/range.md#mergeacross-bool)|Merge the range cells into one region in the worksheet.|1.2|
|[range](resources/range.md)|_Method_ > [unmerge()](resources/range.md#unmerge)|Unmerge the range cells into separate cells.|1.2|
|[rangeFormat](resources/rangeformat.md)|_Property_ > columnWidth|Gets or sets the width of all colums within the range. If the column widths are not uniform, null will be returned.|1.2|
|[rangeFormat](resources/rangeformat.md)|_Property_ > rowHeight|Gets or sets the height of all rows in the range. If the row heights are not uniform null will be returned.|1.2|
|[rangeFormat](resources/rangeformat.md)|_Relationship_ > protection|Returns the format protection object for a range. Read-only.|1.2|
|[rangeFormat](resources/rangeformat.md)|_Method_ > [autofitColumns()](resources/rangeformat.md#autofitcolumns)|Changes the width of the columns of the current range to achieve the best fit, based on the current data in the columns.|1.2|
|[rangeFormat](resources/rangeformat.md)|_Method_ > [autofitRows()](resources/rangeformat.md#autofitrows)|Changes the height of the rows of the current range to achieve the best fit, based on the current data in the columns.|1.2|
|[rangeReference](resources/rangereference.md)|_Property_ > address|Represents the visible rows of the current range.|1.2|
|[rangeSort](resources/rangesort.md)|_Method_ > [apply(fields: SortField[], matchCase: bool, hasHeaders: bool, orientation: string, method: string)](resources/rangesort.md#applyfields-sortfield-matchcase-bool-hasheaders-bool-orientation-string-method-string)|Perform a sort operation.|1.2|
|[selectionChangedEventArgs](resources/selectionchangedeventargs.md)|_Relationship_ > workbook|Gets the workbook object that raised the SelectionChanged event.|1.2, introduced in 1.3|
|[sortField](resources/sortfield.md)|_Property_ > ascending|Represents whether the sorting is done in an ascending fashion.|1.2|
|[sortField](resources/sortfield.md)|_Property_ > color|Represents the color that is the target of the condition if the sorting is on font or cell color.|1.2|
|[sortField](resources/sortfield.md)|_Property_ > dataOption|Represents additional sorting options for this field. Possible values are: Normal, TextAsNumber.|1.2|
|[sortField](resources/sortfield.md)|_Property_ > key|Represents the column (or row, depending on the sort orientation) that the condition is on. Represented as an offset from the first column (or row).|1.2|
|[sortField](resources/sortfield.md)|_Property_ > sortOn|Represents the type of sorting of this condition. Possible values are: Value, CellColor, FontColor, Icon.|1.2|
|[sortField](resources/sortfield.md)|_Relationship_ > icon|Represents the icon that is the target of the condition if the sorting is on the cell's icon.|1.2|
|[table](resources/table.md)|_Relationship_ > sort|Represents the sorting for the table. Read-only.|1.2|
|[table](resources/table.md)|_Relationship_ > worksheet|The worksheet containing the current table. Read-only.|1.2|
|[table](resources/table.md)|_Method_ > [clearFilters()](resources/table.md#clearfilters)|Clears all the filters currently applied on the table.|1.2|
|[table](resources/table.md)|_Method_ > [convertToRange()](resources/table.md#converttorange)|Converts the table into a normal range of cells. All data is preserved.|1.2|
|[table](resources/table.md)|_Method_ > [reapplyFilters()](resources/table.md#reapplyfilters)|Reapplies all the filters currently on the table.|1.2|
|[tableColumn](resources/tablecolumn.md)|_Relationship_ > filter|Retrieve the filter applied to the column. Read-only.|1.2|
|[tableSort](resources/tablesort.md)|_Property_ > matchCase|Represents whether the casing impacted the last sort of the table. Read-only.|1.2|
|[tableSort](resources/tablesort.md)|_Property_ > method|Represents Chinese character ordering method last used to sort the table. Read-only. Possible values are: PinYin, StrokeCount.|1.2|
|[tableSort](resources/tablesort.md)|_Relationship_ > fields|Represents the current conditions used to last sort the table. Read-only.|1.2|
|[tableSort](resources/tablesort.md)|_Method_ > [apply(fields: SortField[], matchCase: bool, method: string)](resources/tablesort.md#applyfields-sortfield-matchcase-bool-method-string)|Perform a sort operation.|1.2|
|[tableSort](resources/tablesort.md)|_Method_ > [clear()](resources/tablesort.md#clear)|Clears the sorting that is currently on the table. While this doesn't modify the table's ordering, it clears the state of the header buttons.|1.2|
|[tableSort](resources/tablesort.md)|_Method_ > [reapply()](resources/tablesort.md#reapply)|Reapplies the current sorting parameters to the table.|1.2|
|[workbook](resources/workbook.md)|_Relationship_ > functions|Represents Excel application instance that contains this workbook. Read-only.|1.2|
|[worksheet](resources/worksheet.md)|_Relationship_ > protection|Returns sheet protection object for a worksheet. Read-only.|1.2|
|[worksheetProtection](resources/worksheetprotection.md)|_Property_ > protected|Indicates if the worksheet is protected. Read-Only. Read-only.|1.2|
|[worksheetProtection](resources/worksheetprotection.md)|_Relationship_ > options|Sheet protection options. Read-only.|1.2|
|[worksheetProtection](resources/worksheetprotection.md)|_Method_ > [protect(options: WorksheetProtectionOptions)](resources/worksheetprotection.md#protectoptions-worksheetprotectionoptions)|Protects a worksheet. Fails if the worksheet has been protected.|1.2|
|[worksheetProtection](resources/worksheetprotection.md)|_Method_ > [unprotect()](resources/worksheetprotection.md#unprotect)|Unprotects a worksheet.|1.2|
|[worksheetProtectionOptions](resources/worksheetprotectionoptions.md)|_Property_ > allowAutoFilter|Represents the worksheet protection option of allowing using auto filter feature.|1.2|
|[worksheetProtectionOptions](resources/worksheetprotectionoptions.md)|_Property_ > allowDeleteColumns|Represents the worksheet protection option of allowing deleting columns.|1.2|
|[worksheetProtectionOptions](resources/worksheetprotectionoptions.md)|_Property_ > allowDeleteRows|Represents the worksheet protection option of allowing deleting rows.|1.2|
|[worksheetProtectionOptions](resources/worksheetprotectionoptions.md)|_Property_ > allowFormatCells|Represents the worksheet protection option of allowing formatting cells.|1.2|
|[worksheetProtectionOptions](resources/worksheetprotectionoptions.md)|_Property_ > allowFormatColumns|Represents the worksheet protection option of allowing formatting columns.|1.2|
|[worksheetProtectionOptions](resources/worksheetprotectionoptions.md)|_Property_ > allowFormatRows|Represents the worksheet protection option of allowing formatting rows.|1.2|
|[worksheetProtectionOptions](resources/worksheetprotectionoptions.md)|_Property_ > allowInsertColumns|Represents the worksheet protection option of allowing inserting columns.|1.2|
|[worksheetProtectionOptions](resources/worksheetprotectionoptions.md)|_Property_ > allowInsertHyperlinks|Represents the worksheet protection option of allowing inserting hyperlinks.|1.2|
|[worksheetProtectionOptions](resources/worksheetprotectionoptions.md)|_Property_ > allowInsertRows|Represents the worksheet protection option of allowing inserting rows.|1.2|
|[worksheetProtectionOptions](resources/worksheetprotectionoptions.md)|_Property_ > allowPivotTables|Represents the worksheet protection option of allowing using PivotTable feature.|1.2|
|[worksheetProtectionOptions](resources/worksheetprotectionoptions.md)|_Property_ > allowSort|Represents the worksheet protection option of allowing using sort feature.|1.2|

## Excel JavaScript API 1.1
Excel JavaScript API 1.1 is the first version of the API. For details about the API,  see the Excel JavaScript API reference topics.  
    
## Additional resources

- [Specify Office hosts and API requirements](../docs/overview/specify-office-hosts-and-api-requirements.md)
- [Office Add-ins XML manifest](https://dev.office.com/docs/add-ins/overview/add-in-manifests)
