---
title: Excel JavaScript API requirement set 1.12
description: 'Details about the ExcelApi 1.12 requirement set.'
ms.date: 04/01/2021
ms.prod: excel
localization_priority: Normal
---

# What's new in Excel JavaScript API 1.12

The ExcelApi 1.12 increased support for formulas in ranges by adding APIs for tracking dynamic arrays and finding a formula's direct precedents. It also added API control of PivotTable filters. Improvements were also made in the comment, culture settings, and custom properties feature areas.

| Feature area | Description | Relevant objects |
|:--- |:--- |:--- |
| [Comment events](../../excel/excel-add-ins-comments.md#comment-events) | Adds events for add, change, and delete to the comment collection.| [CommentCollection](/javascript/api/excel/excel.commentcollection) |
| Date and time [culture settings](../../excel/excel-add-ins-workbooks.md#access-application-culture-settings) | Gives access to additional cultural settings around date and time formatting. | [CultureInfo](/javascript/api/excel/excel.cultureinfo), [NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo) [Application](/javascript/api/excel/excel.application) |
| [Direct precedents](../../excel/excel-add-ins-ranges-precedents.md) | Returns ranges that are used to evaluate a cell's formula.| [Range](/javascript/api/excel/excel.range#getdirectprecedents--) |
| Pivot Filters | Applies value-driven filters to the fields of a PivotTable. | [PivotField](/javascript/api/excel/excel.pivotfield#applyfilter-filter-), [PivotFilters](/javascript/api/excel/excel.pivotFilters) |
| [Range spilling](../../excel/excel-add-ins-ranges-dynamic-arrays.md) | Lets add-ins find ranges associated with [dynamic array](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531) results. | [Range](/javascript/api/excel/excel.range) |
| [Worksheet-level custom properties](../../excel/excel-add-ins-workbooks.md#worksheet-level-custom-properties) | Lets custom properties be scoped to the worksheet-level, in addition to being scoped to the workbook-level. | [WorksheetCustomProperty](/javascript/api/excel/excel.worksheetcustomproperty), [WorksheetCustomPropertyCollection](/javascript/api/excel/excel.worksheetcustompropertycollection)|

## API list

The following table lists the APIs in Excel JavaScript API requirement set 1.12. To view API reference documentation for all APIs supported by Excel JavaScript API requirement set 1.12 or earlier, see [Excel APIs in requirement set 1.12 or earlier](/javascript/api/excel?view=excel-js-1.12&preserve-view=true).

| Class | Fields | Description |
|:---|:---|:---|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[textOrientation](/javascript/api/excel/excel.chartaxistitle#textorientation)|Specifies the angle to which the text is oriented for the chart axis title.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[getDimensionValues(dimension: Excel.ChartSeriesDimension)](/javascript/api/excel/excel.chartseries#getdimensionvalues-dimension-)|Gets the values from a single dimension of the chart series.|
|[Comment](/javascript/api/excel/excel.comment)|[contentType](/javascript/api/excel/excel.comment#contenttype)|Gets the content type of the comment.|
|[CommentAddedEventArgs](/javascript/api/excel/excel.commentaddedeventargs)|[commentDetails](/javascript/api/excel/excel.commentaddedeventargs#commentdetails)|Get the CommentDetail array which contains the comment Id and Ids of its related replies.|
||[source](/javascript/api/excel/excel.commentaddedeventargs#source)|Specifies the source of the event.|
||[type](/javascript/api/excel/excel.commentaddedeventargs#type)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.commentaddedeventargs#worksheetid)|Gets the Id of the worksheet in which the event happened.|
|[CommentChangedEventArgs](/javascript/api/excel/excel.commentchangedeventargs)|[changeType](/javascript/api/excel/excel.commentchangedeventargs#changetype)|Gets the change type that represents how the changed event is triggered.|
||[commentDetails](/javascript/api/excel/excel.commentchangedeventargs#commentdetails)|Get the CommentDetail array which contains the comment Id and Ids of its related replies.|
||[source](/javascript/api/excel/excel.commentchangedeventargs#source)|Specifies the source of the event.|
||[type](/javascript/api/excel/excel.commentchangedeventargs#type)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.commentchangedeventargs#worksheetid)|Gets the Id of the worksheet in which the event happened.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[onAdded](/javascript/api/excel/excel.commentcollection#onadded)|Occurs when the comments are added.|
||[onChanged](/javascript/api/excel/excel.commentcollection#onchanged)|Occurs when comments or replies in a comment collection are changed, including when replies are deleted.|
||[onDeleted](/javascript/api/excel/excel.commentcollection#ondeleted)|Occurs when comments are deleted in the comment collection.|
|[CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs)|[commentDetails](/javascript/api/excel/excel.commentdeletedeventargs#commentdetails)|Get the CommentDetail array which contains the comment Id and Ids of its related replies.|
||[source](/javascript/api/excel/excel.commentdeletedeventargs#source)|Specifies the source of the event.|
||[type](/javascript/api/excel/excel.commentdeletedeventargs#type)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.commentdeletedeventargs#worksheetid)|Gets the Id of the worksheet in which the event happened.|
|[CommentDetail](/javascript/api/excel/excel.commentdetail)|[commentId](/javascript/api/excel/excel.commentdetail#commentid)|Represents the id of comment.|
||[replyIds](/javascript/api/excel/excel.commentdetail#replyids)|Represents the ids of the related replies belong to comment.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[contentType](/javascript/api/excel/excel.commentreply#contenttype)|The content type of the reply.|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[datetimeFormat](/javascript/api/excel/excel.cultureinfo#datetimeformat)|Defines the culturally appropriate format of displaying date and time.|
|[DatetimeFormatInfo](/javascript/api/excel/excel.datetimeformatinfo)|[dateSeparator](/javascript/api/excel/excel.datetimeformatinfo#dateseparator)|Gets the string used as the date separator.|
||[longDatePattern](/javascript/api/excel/excel.datetimeformatinfo#longdatepattern)|Gets the format string for a long date value.|
||[longTimePattern](/javascript/api/excel/excel.datetimeformatinfo#longtimepattern)|Gets the format string for a long time value.|
||[shortDatePattern](/javascript/api/excel/excel.datetimeformatinfo#shortdatepattern)|Gets the format string for a short date value.|
||[timeSeparator](/javascript/api/excel/excel.datetimeformatinfo#timeseparator)|Gets the string used as the time separator.|
|[PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter)|[comparator](/javascript/api/excel/excel.pivotdatefilter#comparator)|The comparator is the static value to which other values are compared.|
||[condition](/javascript/api/excel/excel.pivotdatefilter#condition)|Specifies the condition for the filter, which defines the necessary filtering criteria.|
||[exclusive](/javascript/api/excel/excel.pivotdatefilter#exclusive)|If true, filter *excludes* items that meet criteria.|
||[lowerBound](/javascript/api/excel/excel.pivotdatefilter#lowerbound)|The lower-bound of the range for the `Between` filter condition.|
||[upperBound](/javascript/api/excel/excel.pivotdatefilter#upperbound)|The upper-bound of the range for the `Between` filter condition.|
||[wholeDays](/javascript/api/excel/excel.pivotdatefilter#wholedays)|For `Equals`, `Before`, `After`, and `Between` filter conditions, indicates if comparisons should be made as whole days.|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[applyFilter(filter: Excel.PivotFilters)](/javascript/api/excel/excel.pivotfield#applyfilter-filter-)|Sets one or more of the field's current PivotFilters and applies them to the field.|
||[clearAllFilters()](/javascript/api/excel/excel.pivotfield#clearallfilters--)|Clears all criteria from all of the field's filters.|
||[clearFilter(filterType: Excel.PivotFilterType)](/javascript/api/excel/excel.pivotfield#clearfilter-filtertype-)|Clears all existing criteria from the field's filter of the given type (if one is currently applied).|
||[getFilters()](/javascript/api/excel/excel.pivotfield#getfilters--)|Gets all filters currently applied on the field.|
||[isFiltered(filterType?: Excel.PivotFilterType)](/javascript/api/excel/excel.pivotfield#isfiltered-filtertype-)|Checks if there are any applied filters on the field.|
|[PivotFilters](/javascript/api/excel/excel.pivotfilters)|[dateFilter](/javascript/api/excel/excel.pivotfilters#datefilter)|The PivotField's currently applied date filter.|
||[labelFilter](/javascript/api/excel/excel.pivotfilters#labelfilter)|The PivotField's currently applied label filter.|
||[manualFilter](/javascript/api/excel/excel.pivotfilters#manualfilter)|The PivotField's currently applied manual filter.|
||[valueFilter](/javascript/api/excel/excel.pivotfilters#valuefilter)|The PivotField's currently applied value filter.|
|[PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter)|[comparator](/javascript/api/excel/excel.pivotlabelfilter#comparator)|The comparator is the static value to which other values are compared.|
||[condition](/javascript/api/excel/excel.pivotlabelfilter#condition)|Specifies the condition for the filter, which defines the necessary filtering criteria.|
||[exclusive](/javascript/api/excel/excel.pivotlabelfilter#exclusive)|If true, filter *excludes* items that meet criteria.|
||[lowerBound](/javascript/api/excel/excel.pivotlabelfilter#lowerbound)|The lower-bound of the range for the Between filter condition.|
||[substring](/javascript/api/excel/excel.pivotlabelfilter#substring)|The substring used for `BeginsWith`, `EndsWith`, and `Contains` filter conditions.|
||[upperBound](/javascript/api/excel/excel.pivotlabelfilter#upperbound)|The upper-bound of the range for the Between filter condition.|
|[PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter)|[selectedItems](/javascript/api/excel/excel.pivotmanualfilter#selecteditems)|A list of selected items to manually filter.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[allowMultipleFiltersPerField](/javascript/api/excel/excel.pivottable#allowmultiplefiltersperfield)|Specifies if the PivotTable allows the application of multiple PivotFilters on a given PivotField in the table.|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getCount()](/javascript/api/excel/excel.pivottablescopedcollection#getcount--)|Gets the number of PivotTables in the collection.|
||[getFirst()](/javascript/api/excel/excel.pivottablescopedcollection#getfirst--)|Gets the first PivotTable in the collection.|
||[getItem(key: string)](/javascript/api/excel/excel.pivottablescopedcollection#getitem-key-)|Gets a PivotTable by name.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablescopedcollection#getitemornullobject-name-)|Gets a PivotTable by name.|
||[items](/javascript/api/excel/excel.pivottablescopedcollection#items)|Gets the loaded child items in this collection.|
|[PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter)|[comparator](/javascript/api/excel/excel.pivotvaluefilter#comparator)|The comparator is the static value to which other values are compared.|
||[condition](/javascript/api/excel/excel.pivotvaluefilter#condition)|Specifies the condition for the filter, which defines the necessary filtering criteria.|
||[exclusive](/javascript/api/excel/excel.pivotvaluefilter#exclusive)|If true, filter *excludes* items that meet criteria.|
||[lowerBound](/javascript/api/excel/excel.pivotvaluefilter#lowerbound)|The lower-bound of the range for the `Between` filter condition.|
||[selectionType](/javascript/api/excel/excel.pivotvaluefilter#selectiontype)|Specifies if the filter is for the top/bottom N items, top/bottom N percent, or top/bottom N sum.|
||[threshold](/javascript/api/excel/excel.pivotvaluefilter#threshold)|The "N" threshold number of items, percent, or sum to be filtered for a Top/Bottom filter condition.|
||[upperBound](/javascript/api/excel/excel.pivotvaluefilter#upperbound)|The upper-bound of the range for the `Between` filter condition.|
||[value](/javascript/api/excel/excel.pivotvaluefilter#value)|Name of the chosen "value" in the field by which to filter.|
|[Range](/javascript/api/excel/excel.range)|[getDirectPrecedents()](/javascript/api/excel/excel.range#getdirectprecedents--)|Returns a WorkbookRangeAreas object that represents the range containing all the direct precedents of a cell in same worksheet or in multiple worksheets.|
||[getPivotTables(fullyContained?: boolean)](/javascript/api/excel/excel.range#getpivottables-fullycontained-)|Gets a scoped collection of PivotTables that overlap with the range.|
||[getSpillParent()](/javascript/api/excel/excel.range#getspillparent--)|Gets the range object containing the anchor cell for a cell getting spilled into.|
||[getSpillParentOrNullObject()](/javascript/api/excel/excel.range#getspillparentornullobject--)|Gets the range object containing the anchor cell for a cell getting spilled into.|
||[getSpillingToRange()](/javascript/api/excel/excel.range#getspillingtorange--)|Gets the range object containing the spill range when called on an anchor cell.|
||[getSpillingToRangeOrNullObject()](/javascript/api/excel/excel.range#getspillingtorangeornullobject--)|Gets the range object containing the spill range when called on an anchor cell.|
||[hasSpill](/javascript/api/excel/excel.range#hasspill)|Represents if all cells have a spill border.|
||[numberFormatCategories](/javascript/api/excel/excel.range#numberformatcategories)|Represents the category of number format of each cell.|
||[savedAsArray](/javascript/api/excel/excel.range#savedasarray)|Represents if ALL the cells would be saved as an array formula.|
|[RangeAreasCollection](/javascript/api/excel/excel.rangeareascollection)|[getCount()](/javascript/api/excel/excel.rangeareascollection#getcount--)|Gets the number of RangeAreas objects in this collection.|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangeareascollection#getitemat-index-)|Returns the RangeAreas object based on position in the collection.|
||[items](/javascript/api/excel/excel.rangeareascollection#items)|Gets the loaded child items in this collection.|
|[WorkbookRangeAreas](/javascript/api/excel/excel.workbookrangeareas)|[getRangeAreasBySheet(key: string)](/javascript/api/excel/excel.workbookrangeareas#getrangeareasbysheet-key-)|Returns the `RangeAreas` object based on worksheet id or name in the collection.|
||[getRangeAreasOrNullObjectBySheet(key: string)](/javascript/api/excel/excel.workbookrangeareas#getrangeareasornullobjectbysheet-key-)|Returns the `RangeAreas` object based on worksheet name or id in the collection.|
||[addresses](/javascript/api/excel/excel.workbookrangeareas#addresses)|Returns an array of address in A1-style.|
||[areas](/javascript/api/excel/excel.workbookrangeareas#areas)|Returns the `RangeAreasCollection` object.|
||[ranges](/javascript/api/excel/excel.workbookrangeareas#ranges)|Returns ranges that comprise this object in a `RangeCollection` object.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[customProperties](/javascript/api/excel/excel.worksheet#customproperties)|Gets a collection of worksheet-level custom properties.|
|[WorksheetCustomProperty](/javascript/api/excel/excel.worksheetcustomproperty)|[delete()](/javascript/api/excel/excel.worksheetcustomproperty#delete--)|Deletes the custom property.|
||[key](/javascript/api/excel/excel.worksheetcustomproperty#key)|Gets the key of the custom property.|
||[value](/javascript/api/excel/excel.worksheetcustomproperty#value)|Gets or sets the value of the custom property.|
|[WorksheetCustomPropertyCollection](/javascript/api/excel/excel.worksheetcustompropertycollection)|[add(key: string, value: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#add-key--value-)|Adds a new custom property that maps to the provided key.|
||[getCount()](/javascript/api/excel/excel.worksheetcustompropertycollection#getcount--)|Gets the number of custom properties on this worksheet.|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitem-key-)|Gets a custom property object by its key, which is case-insensitive.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitemornullobject-key-)|Gets a custom property object by its key, which is case-insensitive.|
||[items](/javascript/api/excel/excel.worksheetcustompropertycollection#items)|Gets the loaded child items in this collection.|

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel?view=excel-js-1.12&preserve-view=true)
- [Excel JavaScript API requirement sets](excel-api-requirement-sets.md)
