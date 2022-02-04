---
title: Excel JavaScript API requirement set 1.12
description: 'Details about the ExcelApi 1.12 requirement set.'
ms.date: 04/01/2021
ms.prod: excel
ms.localizationpriority: medium
---

# What's new in Excel JavaScript API 1.12

The ExcelApi 1.12 increased support for formulas in ranges by adding APIs for tracking dynamic arrays and finding a formula's direct precedents. It also added API control of PivotTable filters. Improvements were also made in the comment, culture settings, and custom properties feature areas.

| Feature area | Description | Relevant objects |
|:--- |:--- |:--- |
| [Comment events](../../excel/excel-add-ins-comments.md#comment-events) | Adds events for add, change, and delete to the comment collection.| [CommentCollection](/javascript/api/excel/excel.commentcollection) |
| Date and time [culture settings](../../excel/excel-add-ins-workbooks.md#access-application-culture-settings) | Gives access to additional cultural settings around date and time formatting. | [CultureInfo](/javascript/api/excel/excel.cultureinfo), [NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo) [Application](/javascript/api/excel/excel.application) |
| [Direct precedents](../../excel/excel-add-ins-ranges-precedents.md) | Returns ranges that are used to evaluate a cell's formula.| [Range](/javascript/api/excel/excel.range#getdirectprecedents--) |
| Pivot Filters | Applies value-driven filters to the fields of a PivotTable. | [PivotField](/javascript/api/excel/excel.pivotfield#applyfilter-filter-), [PivotFilters](/javascript/api/excel/excel.pivotfilters) |
| [Range spilling](../../excel/excel-add-ins-ranges-dynamic-arrays.md) | Lets add-ins find ranges associated with [dynamic array](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531) results. | [Range](/javascript/api/excel/excel.range) |
| [Worksheet-level custom properties](../../excel/excel-add-ins-workbooks.md#worksheet-level-custom-properties) | Lets custom properties be scoped to the worksheet-level, in addition to being scoped to the workbook-level. | [WorksheetCustomProperty](/javascript/api/excel/excel.worksheetcustomproperty), [WorksheetCustomPropertyCollection](/javascript/api/excel/excel.worksheetcustompropertycollection)|

## API list

The following table lists the APIs in Excel JavaScript API requirement set 1.12. To view API reference documentation for all APIs supported by Excel JavaScript API requirement set 1.12 or earlier, see [Excel APIs in requirement set 1.12 or earlier](/javascript/api/excel?view=excel-js-1.12&preserve-view=true).

| Class | Fields | Description |
|:---|:---|:---|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[textOrientation](/javascript/api/excel/excel.chartaxistitle#excel-excel-chartaxistitle-textorientation-member)|Specifies the angle to which the text is oriented for the chart axis title.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[getDimensionValues(dimension: Excel.ChartSeriesDimension)](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-getdimensionvalues-member(1))|Gets the values from a single dimension of the chart series.|
|[Comment](/javascript/api/excel/excel.comment)|[contentType](/javascript/api/excel/excel.comment#excel-excel-comment-contenttype-member)|Gets the content type of the comment.|
|[CommentAddedEventArgs](/javascript/api/excel/excel.commentaddedeventargs)|[commentDetails](/javascript/api/excel/excel.commentaddedeventargs#excel-excel-commentaddedeventargs-commentdetails-member)|Gets the `CommentDetail` array that contains the comment ID and IDs of its related replies.|
||[source](/javascript/api/excel/excel.commentaddedeventargs#excel-excel-commentaddedeventargs-source-member)|Specifies the source of the event.|
||[type](/javascript/api/excel/excel.commentaddedeventargs#excel-excel-commentaddedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.commentaddedeventargs#excel-excel-commentaddedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the event happened.|
|[CommentChangedEventArgs](/javascript/api/excel/excel.commentchangedeventargs)|[changeType](/javascript/api/excel/excel.commentchangedeventargs#excel-excel-commentchangedeventargs-changetype-member)|Gets the change type that represents how the changed event is triggered.|
||[commentDetails](/javascript/api/excel/excel.commentchangedeventargs#excel-excel-commentchangedeventargs-commentdetails-member)|Get the `CommentDetail` array which contains the comment ID and IDs of its related replies.|
||[source](/javascript/api/excel/excel.commentchangedeventargs#excel-excel-commentchangedeventargs-source-member)|Specifies the source of the event.|
||[type](/javascript/api/excel/excel.commentchangedeventargs#excel-excel-commentchangedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.commentchangedeventargs#excel-excel-commentchangedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the event happened.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[onAdded](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-onadded-member)|Occurs when the comments are added.|
||[onChanged](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-onchanged-member)|Occurs when comments or replies in a comment collection are changed, including when replies are deleted.|
||[onDeleted](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-ondeleted-member)|Occurs when comments are deleted in the comment collection.|
|[CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs)|[commentDetails](/javascript/api/excel/excel.commentdeletedeventargs#excel-excel-commentdeletedeventargs-commentdetails-member)|Gets the `CommentDetail` array that contains the comment ID and IDs of its related replies.|
||[source](/javascript/api/excel/excel.commentdeletedeventargs#excel-excel-commentdeletedeventargs-source-member)|Specifies the source of the event.|
||[type](/javascript/api/excel/excel.commentdeletedeventargs#excel-excel-commentdeletedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.commentdeletedeventargs#excel-excel-commentdeletedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the event happened.|
|[CommentDetail](/javascript/api/excel/excel.commentdetail)|[commentId](/javascript/api/excel/excel.commentdetail#excel-excel-commentdetail-commentid-member)|Represents the ID of the comment.|
||[replyIds](/javascript/api/excel/excel.commentdetail#excel-excel-commentdetail-replyids-member)|Represents the IDs of the related replies that belong to the comment.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[contentType](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-contenttype-member)|The content type of the reply.|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[datetimeFormat](/javascript/api/excel/excel.cultureinfo#excel-excel-cultureinfo-datetimeformat-member)|Defines the culturally appropriate format of displaying date and time.|
|[DatetimeFormatInfo](/javascript/api/excel/excel.datetimeformatinfo)|[dateSeparator](/javascript/api/excel/excel.datetimeformatinfo#excel-excel-datetimeformatinfo-dateseparator-member)|Gets the string used as the date separator.|
||[longDatePattern](/javascript/api/excel/excel.datetimeformatinfo#excel-excel-datetimeformatinfo-longdatepattern-member)|Gets the format string for a long date value.|
||[longTimePattern](/javascript/api/excel/excel.datetimeformatinfo#excel-excel-datetimeformatinfo-longtimepattern-member)|Gets the format string for a long time value.|
||[shortDatePattern](/javascript/api/excel/excel.datetimeformatinfo#excel-excel-datetimeformatinfo-shortdatepattern-member)|Gets the format string for a short date value.|
||[timeSeparator](/javascript/api/excel/excel.datetimeformatinfo#excel-excel-datetimeformatinfo-timeseparator-member)|Gets the string used as the time separator.|
|[PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter)|[comparator](/javascript/api/excel/excel.pivotdatefilter#excel-excel-pivotdatefilter-comparator-member)|The comparator is the static value to which other values are compared.|
||[condition](/javascript/api/excel/excel.pivotdatefilter#excel-excel-pivotdatefilter-condition-member)|Specifies the condition for the filter, which defines the necessary filtering criteria.|
||[exclusive](/javascript/api/excel/excel.pivotdatefilter#excel-excel-pivotdatefilter-exclusive-member)|If `true`, filter *excludes* items that meet criteria.|
||[lowerBound](/javascript/api/excel/excel.pivotdatefilter#excel-excel-pivotdatefilter-lowerbound-member)|The lower-bound of the range for the `between` filter condition.|
||[upperBound](/javascript/api/excel/excel.pivotdatefilter#excel-excel-pivotdatefilter-upperbound-member)|The upper-bound of the range for the `between` filter condition.|
||[wholeDays](/javascript/api/excel/excel.pivotdatefilter#excel-excel-pivotdatefilter-wholedays-member)|For `equals`, `before`, `after`, and `between` filter conditions, indicates if comparisons should be made as whole days.|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[applyFilter(filter: Excel.PivotFilters)](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-applyfilter-member(1))|Sets one or more of the field's current PivotFilters and applies them to the field.|
||[clearAllFilters()](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-clearallfilters-member(1))|Clears all criteria from all of the field's filters.|
||[clearFilter(filterType: Excel.PivotFilterType)](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-clearfilter-member(1))|Clears all existing criteria from the field's filter of the given type (if one is currently applied).|
||[getFilters()](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-getfilters-member(1))|Gets all filters currently applied on the field.|
||[isFiltered(filterType?: Excel.PivotFilterType)](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-isfiltered-member(1))|Checks if there are any applied filters on the field.|
|[PivotFilters](/javascript/api/excel/excel.pivotfilters)|[dateFilter](/javascript/api/excel/excel.pivotfilters#excel-excel-pivotfilters-datefilter-member)|The PivotField's currently applied date filter.|
||[labelFilter](/javascript/api/excel/excel.pivotfilters#excel-excel-pivotfilters-labelfilter-member)|The PivotField's currently applied label filter.|
||[manualFilter](/javascript/api/excel/excel.pivotfilters#excel-excel-pivotfilters-manualfilter-member)|The PivotField's currently applied manual filter.|
||[valueFilter](/javascript/api/excel/excel.pivotfilters#excel-excel-pivotfilters-valuefilter-member)|The PivotField's currently applied value filter.|
|[PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter)|[comparator](/javascript/api/excel/excel.pivotlabelfilter#excel-excel-pivotlabelfilter-comparator-member)|The comparator is the static value to which other values are compared.|
||[condition](/javascript/api/excel/excel.pivotlabelfilter#excel-excel-pivotlabelfilter-condition-member)|Specifies the condition for the filter, which defines the necessary filtering criteria.|
||[exclusive](/javascript/api/excel/excel.pivotlabelfilter#excel-excel-pivotlabelfilter-exclusive-member)|If `true`, filter *excludes* items that meet criteria.|
||[lowerBound](/javascript/api/excel/excel.pivotlabelfilter#excel-excel-pivotlabelfilter-lowerbound-member)|The lower-bound of the range for the `between` filter condition.|
||[substring](/javascript/api/excel/excel.pivotlabelfilter#excel-excel-pivotlabelfilter-substring-member)|The substring used for `beginsWith`, `endsWith`, and `contains` filter conditions.|
||[upperBound](/javascript/api/excel/excel.pivotlabelfilter#excel-excel-pivotlabelfilter-upperbound-member)|The upper-bound of the range for the `between` filter condition.|
|[PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter)|[selectedItems](/javascript/api/excel/excel.pivotmanualfilter#excel-excel-pivotmanualfilter-selecteditems-member)|A list of selected items to manually filter.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[allowMultipleFiltersPerField](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-allowmultiplefiltersperfield-member)|Specifies if the PivotTable allows the application of multiple PivotFilters on a given PivotField in the table.|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getCount()](/javascript/api/excel/excel.pivottablescopedcollection#excel-excel-pivottablescopedcollection-getcount-member(1))|Gets the number of PivotTables in the collection.|
||[getFirst()](/javascript/api/excel/excel.pivottablescopedcollection#excel-excel-pivottablescopedcollection-getfirst-member(1))|Gets the first PivotTable in the collection.|
||[getItem(key: string)](/javascript/api/excel/excel.pivottablescopedcollection#excel-excel-pivottablescopedcollection-getitem-member(1))|Gets a PivotTable by name.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablescopedcollection#excel-excel-pivottablescopedcollection-getitemornullobject-member(1))|Gets a PivotTable by name.|
||[items](/javascript/api/excel/excel.pivottablescopedcollection#excel-excel-pivottablescopedcollection-items-member)|Gets the loaded child items in this collection.|
|[PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter)|[comparator](/javascript/api/excel/excel.pivotvaluefilter#excel-excel-pivotvaluefilter-comparator-member)|The comparator is the static value to which other values are compared.|
||[condition](/javascript/api/excel/excel.pivotvaluefilter#excel-excel-pivotvaluefilter-condition-member)|Specifies the condition for the filter, which defines the necessary filtering criteria.|
||[exclusive](/javascript/api/excel/excel.pivotvaluefilter#excel-excel-pivotvaluefilter-exclusive-member)|If `true`, filter *excludes* items that meet criteria.|
||[lowerBound](/javascript/api/excel/excel.pivotvaluefilter#excel-excel-pivotvaluefilter-lowerbound-member)|The lower-bound of the range for the `between` filter condition.|
||[selectionType](/javascript/api/excel/excel.pivotvaluefilter#excel-excel-pivotvaluefilter-selectiontype-member)|Specifies if the filter is for the top/bottom N items, top/bottom N percent, or top/bottom N sum.|
||[threshold](/javascript/api/excel/excel.pivotvaluefilter#excel-excel-pivotvaluefilter-threshold-member)|The "N" threshold number of items, percent, or sum to be filtered for a top/bottom filter condition.|
||[upperBound](/javascript/api/excel/excel.pivotvaluefilter#excel-excel-pivotvaluefilter-upperbound-member)|The upper-bound of the range for the `between` filter condition.|
||[value](/javascript/api/excel/excel.pivotvaluefilter#excel-excel-pivotvaluefilter-value-member)|Name of the chosen "value" in the field by which to filter.|
|[Range](/javascript/api/excel/excel.range)|[getDirectPrecedents()](/javascript/api/excel/excel.range#excel-excel-range-getdirectprecedents-member(1))|Returns a `WorkbookRangeAreas` object that represents the range containing all the direct precedents of a cell in the same worksheet or in multiple worksheets.|
||[getPivotTables(fullyContained?: boolean)](/javascript/api/excel/excel.range#excel-excel-range-getpivottables-member(1))|Gets a scoped collection of PivotTables that overlap with the range.|
||[getSpillParent()](/javascript/api/excel/excel.range#excel-excel-range-getspillparent-member(1))|Gets the range object containing the anchor cell for a cell getting spilled into.|
||[getSpillParentOrNullObject()](/javascript/api/excel/excel.range#excel-excel-range-getspillparentornullobject-member(1))|Gets the range object containing the anchor cell for the cell getting spilled into.|
||[getSpillingToRange()](/javascript/api/excel/excel.range#excel-excel-range-getspillingtorange-member(1))|Gets the range object containing the spill range when called on an anchor cell.|
||[getSpillingToRangeOrNullObject()](/javascript/api/excel/excel.range#excel-excel-range-getspillingtorangeornullobject-member(1))|Gets the range object containing the spill range when called on an anchor cell.|
||[hasSpill](/javascript/api/excel/excel.range#excel-excel-range-hasspill-member)|Represents if all cells have a spill border.|
||[numberFormatCategories](/javascript/api/excel/excel.range#excel-excel-range-numberformatcategories-member)|Represents the category of number format of each cell.|
||[savedAsArray](/javascript/api/excel/excel.range#excel-excel-range-savedasarray-member)|Represents if all the cells would be saved as an array formula.|
|[RangeAreasCollection](/javascript/api/excel/excel.rangeareascollection)|[getCount()](/javascript/api/excel/excel.rangeareascollection#excel-excel-rangeareascollection-getcount-member(1))|Gets the number of `RangeAreas` objects in this collection.|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangeareascollection#excel-excel-rangeareascollection-getitemat-member(1))|Returns the `RangeAreas` object based on position in the collection.|
||[items](/javascript/api/excel/excel.rangeareascollection#excel-excel-rangeareascollection-items-member)|Gets the loaded child items in this collection.|
|[WorkbookRangeAreas](/javascript/api/excel/excel.workbookrangeareas)|[addresses](/javascript/api/excel/excel.workbookrangeareas#excel-excel-workbookrangeareas-addresses-member)|Returns an array of addresses in A1-style.|
||[areas](/javascript/api/excel/excel.workbookrangeareas#excel-excel-workbookrangeareas-areas-member)|Returns the `RangeAreasCollection` object.|
||[getRangeAreasBySheet(key: string)](/javascript/api/excel/excel.workbookrangeareas#excel-excel-workbookrangeareas-getrangeareasbysheet-member(1))|Returns the `RangeAreas` object based on worksheet ID or name in the collection.|
||[getRangeAreasOrNullObjectBySheet(key: string)](/javascript/api/excel/excel.workbookrangeareas#excel-excel-workbookrangeareas-getrangeareasornullobjectbysheet-member(1))|Returns the `RangeAreas` object based on worksheet name or ID in the collection.|
||[ranges](/javascript/api/excel/excel.workbookrangeareas#excel-excel-workbookrangeareas-ranges-member)|Returns ranges that comprise this object in a `RangeCollection` object.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[customProperties](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-customproperties-member)|Gets a collection of worksheet-level custom properties.|
|[WorksheetCustomProperty](/javascript/api/excel/excel.worksheetcustomproperty)|[delete()](/javascript/api/excel/excel.worksheetcustomproperty#excel-excel-worksheetcustomproperty-delete-member(1))|Deletes the custom property.|
||[key](/javascript/api/excel/excel.worksheetcustomproperty#excel-excel-worksheetcustomproperty-key-member)|Gets the key of the custom property.|
||[value](/javascript/api/excel/excel.worksheetcustomproperty#excel-excel-worksheetcustomproperty-value-member)|Gets or sets the value of the custom property.|
|[WorksheetCustomPropertyCollection](/javascript/api/excel/excel.worksheetcustompropertycollection)|[add(key: string, value: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#excel-excel-worksheetcustompropertycollection-add-member(1))|Adds a new custom property that maps to the provided key.|
||[getCount()](/javascript/api/excel/excel.worksheetcustompropertycollection#excel-excel-worksheetcustompropertycollection-getcount-member(1))|Gets the number of custom properties on this worksheet.|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#excel-excel-worksheetcustompropertycollection-getitem-member(1))|Gets a custom property object by its key, which is case-insensitive.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#excel-excel-worksheetcustompropertycollection-getitemornullobject-member(1))|Gets a custom property object by its key, which is case-insensitive.|
||[items](/javascript/api/excel/excel.worksheetcustompropertycollection#excel-excel-worksheetcustompropertycollection-items-member)|Gets the loaded child items in this collection.|

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel?view=excel-js-1.12&preserve-view=true)
- [Excel JavaScript API requirement sets](excel-api-requirement-sets.md)
