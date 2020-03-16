---
title: Excel JavaScript preview APIs
description: 'Details about upcoming Excel JavaScript APIs'
ms.date: 01/02/2020
ms.prod: excel
localization_priority: Normal
---

# Excel JavaScript preview APIs

New Excel JavaScript APIs are first introduced in "preview" and later become part of a specific, numbered requirement set after sufficient testing occurs and user feedback is acquired.

The first table provides a concise summary of the APIs, while the subsequent table gives a detailed list.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| Feature area | Description | Relevant objects |
|:--- |:--- |:--- |
| Culture settings | Gets cultural system settings for the workbook, such as number formatting. | [CultureInfo](/javascript/api/excel/excel.cultureinfo), [NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo) [Application](/javascript/api/excel/excel.application) |
| [Insert workbook](../../excel/excel-add-ins-workbooks.md#insert-a-copy-of-an-existing-workbook-into-the-current-one-preview) | Insert one workbook into another.  | [Workbook](/javascript/api/excel/excel.worksheetcollection) |
| Pivot Filters | Applies value-driven filters to the fields of a PivotTable. | [PivotField](/javascript/api/excel/excel.pivotfield#applyfilter-filter-), [PivotFilters](/javascript/api/excel/excel.pivotFilters) |
| Workbook [Save](../../excel/excel-add-ins-workbooks.md#save-the-workbook-preview) and [Close](../../excel/excel-add-ins-workbooks.md#close-the-workbook-preview) | Save and close workbooks.  | [Workbook](/javascript/api/excel/excel.workbook) |

## API list

The following table lists the Excel JavaScript APIs currently in preview. To see a complete list of all Excel JavaScript APIs (including preview APIs and previously released APIs), see [all Excel JavaScript APIs](/javascript/api/excel?view=excel-js-preview).

| Class | Fields | Description |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[cultureInfo](/javascript/api/excel/excel.application#cultureinfo)|Provides information based on current system culture settings. This includes the culture names, number formatting, and other culturally dependent settings.|
||[decimalSeparator](/javascript/api/excel/excel.application#decimalseparator)|Gets the string used as the decimal separator for numeric values. This is based on Excel's local settings.|
||[thousandsSeparator](/javascript/api/excel/excel.application#thousandsseparator)|Gets the string used to separate groups of digits to the left of the decimal for numeric values. This is based on Excel's local settings.|
||[useSystemSeparators](/javascript/api/excel/excel.application#usesystemseparators)|Specifies whether the system separators of Microsoft Excel are enabled.|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[textOrientation](/javascript/api/excel/excel.chartaxistitle#textorientation)|Represents the angle to which the text is oriented for the chart axis title. The value should either be an integer from -90 to 90 or the integer 180 for vertically-oriented text.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[getDimensionValues(dimension: Excel.ChartSeriesDimension)](/javascript/api/excel/excel.chartseries#getdimensionvalues-dimension-)|Gets the values from a single dimension of the chart series. These could be either category values or data values, depending on the dimension specified and how the data is mapped for the chart series.|
|[Comment](/javascript/api/excel/excel.comment)|[contentType](/javascript/api/excel/excel.comment#contenttype)|Gets the content type of the comment.|
||[resolved](/javascript/api/excel/excel.comment#resolved)|Gets or sets the comment thread status. A value of "true" means the comment thread is in the resolved state.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[contentType](/javascript/api/excel/excel.commentreply#contenttype)|Gets the content type of the comment.|
||[resolved](/javascript/api/excel/excel.commentreply#resolved)|Gets or sets the comment reply status. A value of "true" means the comment reply is in the resolved state.|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[datetimeFormat](/javascript/api/excel/excel.cultureinfo#datetimeformat)|Defines the culturally appropriate format of displaying date and time. This is based on current system culture settings.|
||[name](/javascript/api/excel/excel.cultureinfo#name)|Gets the culture name in the format languagecode2-country/regioncode2 (e.g. "zh-cn" or "en-us"). This is based on current system settings.|
||[numberFormat](/javascript/api/excel/excel.cultureinfo#numberformat)|Defines the culturally appropriate format of displaying numbers. This is based on current system culture settings.|
|[DatetimeFormatInfo](/javascript/api/excel/excel.datetimeformatinfo)|[dateSeparator](/javascript/api/excel/excel.datetimeformatinfo#dateseparator)|Gets the string used as the date separator. This is based on current system settings.|
||[longDatePattern](/javascript/api/excel/excel.datetimeformatinfo#longdatepattern)|Gets the format string for a long date value. This is based on current system settings.|
||[longTimePattern](/javascript/api/excel/excel.datetimeformatinfo#longtimepattern)|Gets the format string for a long time value. This is based on current system settings.|
||[shortDatePattern](/javascript/api/excel/excel.datetimeformatinfo#shortdatepattern)|Gets the format string for a short date value. This is based on current system settings.|
||[timeSeparator](/javascript/api/excel/excel.datetimeformatinfo#timeseparator)|Gets the string used as the time separator. This is based on current system settings.|
|[NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo)|[numberDecimalSeparator](/javascript/api/excel/excel.numberformatinfo#numberdecimalseparator)|Gets the string used as the decimal separator for numeric values. This is based on current system settings.|
||[numberGroupSeparator](/javascript/api/excel/excel.numberformatinfo#numbergroupseparator)|Gets the string used to separate groups of digits to the left of the decimal for numeric values. This is based on current system settings.|
|[PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter)|[comparator](/javascript/api/excel/excel.pivotdatefilter#comparator)|The comparator is the static value to which other values are compared. The type of comparison is defined by the condition.|
||[condition](/javascript/api/excel/excel.pivotdatefilter#condition)|Indicates the condition for the filter, which defines the necessary filtering criteria.|
||[exclusive](/javascript/api/excel/excel.pivotdatefilter#exclusive)|If true, filter *excludes* items that meet criteria. The default is false (filter to include items that meet criteria).|
||[lowerBound](/javascript/api/excel/excel.pivotdatefilter#lowerbound)|The lower-bound of the range for the `Between` filter condition.|
||[upperBound](/javascript/api/excel/excel.pivotdatefilter#upperbound)|The upper-bound of the range for the `Between` filter condition.|
||[wholeDays](/javascript/api/excel/excel.pivotdatefilter#wholedays)|For `Equals`, `Before`, `After`, and `Between` filter conditions, indicates if comparisons should be made as whole days.|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[applyFilter(filter: PivotValueFilter \| PivotLabelFilter \| PivotManualFilter \| PivotDateFilter \| PivotFilters)](/javascript/api/excel/excel.pivotfield#applyfilter-filter-)|Sets one or multiple of the field's current PivotFilters and applies them to the field.|
||[clearAllFilters()](/javascript/api/excel/excel.pivotfield#clearallfilters--)|Clears all criteria from all of the field's filters. This removes any active filtering on the field.|
||[clearFilter(filterType: Excel.PivotFilterType)](/javascript/api/excel/excel.pivotfield#clearfilter-filtertype-)|Clears all existing criteria from the field's filter of the given type (if one is currently applied).|
||[getFilters()](/javascript/api/excel/excel.pivotfield#getfilters--)|Gets all filters currently applied on the field.|
||[isFiltered(filterType?: Excel.PivotFilterType)](/javascript/api/excel/excel.pivotfield#isfiltered-filtertype-)|Checks if there are any applied filters on the field.|
|[PivotFilters](/javascript/api/excel/excel.pivotfilters)|[dateFilter](/javascript/api/excel/excel.pivotfilters#datefilter)|The PivotField's currently applied date filter. Null if none is applied.|
||[labelFilter](/javascript/api/excel/excel.pivotfilters#labelfilter)|The PivotField's currently applied label filter. Null if none is applied.|
||[manualFilter](/javascript/api/excel/excel.pivotfilters#manualfilter)|The PivotField's currently applied manual filter. Null if none is applied.|
||[valueFilter](/javascript/api/excel/excel.pivotfilters#valuefilter)|The PivotField's currently applied value filter. Null if none is applied.|
|[PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter)|[comparator](/javascript/api/excel/excel.pivotlabelfilter#comparator)|The comparator is the static value to which other values are compared. The type of comparison is defined by the condition.|
||[condition](/javascript/api/excel/excel.pivotlabelfilter#condition)|Indicates the condition for the filter, which defines the necessary filtering criteria.|
||[exclusive](/javascript/api/excel/excel.pivotlabelfilter#exclusive)|If true, filter *excludes* items that meet criteria. The default is false (filter to include items that meet criteria).|
||[lowerBound](/javascript/api/excel/excel.pivotlabelfilter#lowerbound)|The lower-bound of the range for the Between filter condition.|
||[substring](/javascript/api/excel/excel.pivotlabelfilter#substring)|The substring used for `BeginsWith`, `EndsWith`, and `Contains` filter conditions.|
||[upperBound](/javascript/api/excel/excel.pivotlabelfilter#upperbound)|The upper-bound of the range for the Between filter condition.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|Gets a unique cell in the PivotTable based on a data hierarchy and the row and column items of their respective hierarchies. The returned cell is the intersection of the given row and column that contains the data from the given hierarchy. This method is the inverse of calling getPivotItems and getDataHierarchy on a particular cell.|
||[pivotStyle](/javascript/api/excel/excel.pivotlayout#pivotstyle)|The style applied to the PivotTable.|
||[setStyle(style: string \| PivotTableStyle \| BuiltInPivotTableStyle)](/javascript/api/excel/excel.pivotlayout#setstyle-style-)|Sets the style applied to the PivotTable.|
|[PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter)|[selectedItems](/javascript/api/excel/excel.pivotmanualfilter#selecteditems)|A list of selected items to manually filter. These must be existing and valid items from the chosen field.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[allowMultipleFiltersPerField](/javascript/api/excel/excel.pivottable#allowmultiplefiltersperfield)|Specifies whether the PivotTable allows the application of multiple PivotFilters on a given PivotField in the table.|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getCount()](/javascript/api/excel/excel.pivottablescopedcollection#getcount--)|Gets the number of PivotTables in the collection.|
||[getFirst()](/javascript/api/excel/excel.pivottablescopedcollection#getfirst--)|Gets the first PivotTable in the collection. The PivotTables in the collection are sorted top to bottom and left to right, such that top-left table is the first PivotTable in the collection.|
||[getItem(key: string)](/javascript/api/excel/excel.pivottablescopedcollection#getitem-key-)|Gets a PivotTable by name.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablescopedcollection#getitemornullobject-name-)|Gets a PivotTable by name. If the PivotTable does not exist, will return a null object.|
||[items](/javascript/api/excel/excel.pivottablescopedcollection#items)|Gets the loaded child items in this collection.|
|[PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter)|[comparator](/javascript/api/excel/excel.pivotvaluefilter#comparator)|The comparator is the static value to which other values are compared. The type of comparison is defined by the condition.|
||[condition](/javascript/api/excel/excel.pivotvaluefilter#condition)|Indicates the condition for the filter, which defines the necessary filtering criteria.|
||[exclusive](/javascript/api/excel/excel.pivotvaluefilter#exclusive)|If true, filter *excludes* items that meet criteria. The default is false (filter to include items that meet criteria).|
||[lowerBound](/javascript/api/excel/excel.pivotvaluefilter#lowerbound)|The lower-bound of the range for the `Between` filter condition.|
||[selectionType](/javascript/api/excel/excel.pivotvaluefilter#selectiontype)|Indicates whether the filter is for the top/bottom N items, top/bottom N percent, or top/bottom N sum.|
||[threshold](/javascript/api/excel/excel.pivotvaluefilter#threshold)|The "N" threshold number of items, percent, or sum to be filtered for a Top/Bottom filter condition.|
||[upperBound](/javascript/api/excel/excel.pivotvaluefilter#upperbound)|The upper-bound of the range for the `Between` filter condition.|
||[value](/javascript/api/excel/excel.pivotvaluefilter#value)|Name of the chosen "value" in the field by which to filter.|
|[Range](/javascript/api/excel/excel.range)|[getPivotTables(fullyContained?: boolean)](/javascript/api/excel/excel.range#getpivottables-fullycontained-)|Gets a scoped collection of PivotTables that overlap with the range.|
||[getSpillParent()](/javascript/api/excel/excel.range#getspillparent--)|Gets the range object containing the anchor cell for a cell getting spilled into. Fails if applied to a range with more than one cell. Read-only.|
||[getSpillParentOrNullObject()](/javascript/api/excel/excel.range#getspillparentornullobject--)|Gets the range object containing the anchor cell for a cell getting spilled into. Read-only.|
||[getSpillingToRange()](/javascript/api/excel/excel.range#getspillingtorange--)|Gets the range object containing the spill range when called on an anchor cell. Fails if applied to a range with more than one cell. Read-only.|
||[getSpillingToRangeOrNullObject()](/javascript/api/excel/excel.range#getspillingtorangeornullobject--)|Gets the range object containing the spill range when called on an anchor cell. Read-only.|
||[hasSpill](/javascript/api/excel/excel.range#hasspill)|Represents if all cells have a spill border.|
||[numberFormatCategories](/javascript/api/excel/excel.range#numberformatcategories)|Represents the category of number format of each cell. Read-only.|
||[savedAsArray](/javascript/api/excel/excel.range#savedasarray)|Represents if ALL the cells would be saved as an array formula.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|Creates a scalable vector graphic (SVG) from an XML string and adds it to the worksheet. Returns a Shape object that represents the new image.|
|[Slicer](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|Represents the slicer name used in the formula.|
||[slicerStyle](/javascript/api/excel/excel.slicer#slicerstyle)|The style applied to the Slicer.|
||[setStyle(style: string \| PivotTableStyle \| BuiltInSlicerStyle)](/javascript/api/excel/excel.slicer#setstyle-style-)|Sets the style applied to the slicer.|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|Changes the table to use the default table style.|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|Occurs when filter is applied on a specific table.|
||[tableStyle](/javascript/api/excel/excel.table#tablestyle)|The style applied to the Table.|
||[setStyle(style: string \| PivotTableStyle \| BuiltInTableStyle)](/javascript/api/excel/excel.table#setstyle-style-)|Sets the style applied to the slicer.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|Occurs when filter is applied on any table in a workbook, or a worksheet.|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|Gets the id of the table in which the filter is applied.|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|Gets the type of the event. See Excel.EventType for details.|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|Gets the id of the worksheet which contains the table.|
|[Workbook](/javascript/api/excel/excel.workbook)|[close(closeBehavior?: Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#close-closebehavior-)|Close current workbook.|
||[save(saveBehavior?: Excel.SaveBehavior)](/javascript/api/excel/excel.workbook#save-savebehavior-)|Save current workbook.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|True if the workbook uses the 1904 date system.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[customProperties](/javascript/api/excel/excel.worksheet#customproperties)|Returns a collection of worksheet-level custom properties.|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|Occurs when filter is applied on a specific worksheet.|
||[onRowHiddenChanged](/javascript/api/excel/excel.worksheet#onrowhiddenchanged)|Occurs when the hidden state of one or more rows has changed on a specific worksheet.|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[address](/javascript/api/excel/excel.worksheetcalculatedeventargs#address)|The address of the range that completed calculation.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|Inserts the specified worksheets of a workbook into the current workbook.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|Occurs when any worksheet's filter is applied in the workbook.|
||[onRowHiddenChanged](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged)|Occurs when the hidden state of one or more rows has changed on a specific worksheet.|
|[WorksheetCustomProperty](/javascript/api/excel/excel.worksheetcustomproperty)|[key](/javascript/api/excel/excel.worksheetcustomproperty#key)|Gets the key of the custom property. Read only.|
||[value](/javascript/api/excel/excel.worksheetcustomproperty#value)|Gets the value of the custom property. Read only.|
|[WorksheetCustomPropertyCollection](/javascript/api/excel/excel.worksheetcustompropertycollection)|[getCount()](/javascript/api/excel/excel.worksheetcustompropertycollection#getcount--)|Gets the number of custom properties on this worksheet.|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitem-key-)|Gets a custom property object by its key, which is case-insensitive. Throws if the custom property does not exist.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitemornullobject-key-)|Gets a custom property object by its key, which is case-insensitive. Returns a null object if the custom property does not exist.|
||[items](/javascript/api/excel/excel.worksheetcustompropertycollection#items)|Gets the loaded child items in this collection.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|Gets the type of the event. See Excel.EventType for details.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|Gets the id of the worksheet in which the filter is applied.|
|[WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs)|[address](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#address)|Gets the range address that represents the changed area of a specific worksheet.|
||[changeType](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#changetype)|Gets the type of change that represents how the event was triggered. See `Excel.RowHiddenChangeType` for details.|
||[source](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#source)|Gets the source of the event. See Excel.EventSource for details.|
||[type](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#type)|Gets the type of the event. See Excel.EventType for details.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#worksheetid)|Gets the id of the worksheet in which the data changed.|

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel?view=excel-js-preview)
- [Excel JavaScript API requirement sets](./excel-api-requirement-sets.md)
