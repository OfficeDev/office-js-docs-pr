---
title: Excel JavaScript preview APIs
description: 'Details about upcoming Excel JavaScript APIs.'
ms.date: 07/23/2021
ms.prod: excel
ms.localizationpriority: medium
---

# Excel JavaScript preview APIs

New Excel JavaScript APIs are first introduced in "preview" and later become part of a specific, numbered requirement set after sufficient testing occurs and user feedback is acquired.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

The following table provides a concise summary of the APIs, while the subsequent [API list](#api-list) table gives a detailed list.

| Feature area | Description | Relevant objects |
|:--- |:--- |:--- |
| Chart data tables | Control appearance, formatting, and visibility of data tables on charts. | [Chart](/javascript/api/excel/excel.chart), [ChartDataTable](/javascript/api/excel/excel.chartdatatable), [ChartDataTableFormat](/javascript/api/excel/excel.chartdatatableformat) |
| Document tasks | Turn comments into tasks assigned to users. | [DocumentTask](/javascript/api/excel/excel.documenttask) |
| Identities | Manage user identities, including display name and email address. | [Identity](/javascript/api/excel/excel.identity), [IdentityCollection](/javascript/api/excel/excel.identitycollection), [IdentityEntity](/javascript/api/excel/excel.identityentity) |
| Linked data types | Adds support for data types connected to Excel from external sources. | [LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|
| Linked workbooks | Manage links between workbooks, including support for refreshing and breaking workbook links. | [LinkedWorkbook](/javascript/api/excel/excel.linkedworkbook), [LinkedWorkbookCollection](/javascript/api/excel/excel.linkedworkbookcollection) |
| Table styles | Provides control for font, border, fill color, and other aspects of table styles. | [Table](/javascript/api/excel/excel.table), [PivotTable](/javascript/api/excel/excel.pivottable), [Slicer](/javascript/api/excel/excel.slicer) |
| Queries | Retrieve query attributes like name, refresh date, and query count. | [Query](/javascript/api/excel/excel.query), [QueryCollection](/javascript/api/excel/excel.querycollection)|

## API list

The following table lists the Excel JavaScript APIs currently in preview. For a complete list of all Excel JavaScript APIs (including preview APIs and previously released APIs), see [all Excel JavaScript APIs](/javascript/api/excel?view=excel-js-preview&preserve-view=true).

| Class | Fields | Description |
|:---|:---|:---|
|[ChangeDirectionState](/javascript/api/excel/excel.changedirectionstate)|[deleteShiftDirection](/javascript/api/excel/excel.changedirectionstate#deleteShiftDirection)|Represents the direction (such as up or to the left) that the remaining cells will shift when a cell or cells are deleted.|
||[insertShiftDirection](/javascript/api/excel/excel.changedirectionstate#insertShiftDirection)|Represents the direction (such as down or to the right) that the existing cells will shift when a new cell or cells are inserted.|
|[Chart](/javascript/api/excel/excel.chart)|[getDataTable()](/javascript/api/excel/excel.chart#getDataTable__)|Gets the data table on the chart.|
||[getDataTableOrNullObject()](/javascript/api/excel/excel.chart#getDataTableOrNullObject__)|Gets the data table on the chart.|
|[ChartDataTable](/javascript/api/excel/excel.chartdatatable)|[format](/javascript/api/excel/excel.chartdatatable#format)|Represents the format of a chart data table, which includes fill, font, and border format.|
||[showHorizontalBorder](/javascript/api/excel/excel.chartdatatable#showHorizontalBorder)|Specifies whether to display horizontal border of the data table.|
||[showLegendKey](/javascript/api/excel/excel.chartdatatable#showLegendKey)|Specifies whether to show legendkey of the data table.|
||[showOutlineBorder](/javascript/api/excel/excel.chartdatatable#showOutlineBorder)|Specifies whether to display outline border of the data table.|
||[showVerticalBorder](/javascript/api/excel/excel.chartdatatable#showVerticalBorder)|Specifies whether to display vertical border of the data table.|
||[visible](/javascript/api/excel/excel.chartdatatable#visible)|Specifies whether to show data table of the chart.|
|[ChartDataTableFormat](/javascript/api/excel/excel.chartdatatableformat)|[border](/javascript/api/excel/excel.chartdatatableformat#border)|Represents the border format of chart data table, which includes color, line style, and weight.|
||[fill](/javascript/api/excel/excel.chartdatatableformat#fill)|Represents the fill format of an object, which includes background formatting information.|
||[font](/javascript/api/excel/excel.chartdatatableformat#font)|Represents the font attributes (such as font name, font size, and color) for the current object.|
|[Comment](/javascript/api/excel/excel.comment)|[assignTask(assignee: Identity)](/javascript/api/excel/excel.comment#assignTask_assignee_)|Assigns the task attached to the comment to the given user as an assignee.|
||[getTask()](/javascript/api/excel/excel.comment#getTask__)|Gets the task associated with this comment.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.comment#getTaskOrNullObject__)|Gets the task associated with this comment.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[getItemOrNullObject(commentId: string)](/javascript/api/excel/excel.commentcollection#getItemOrNullObject_commentId_)|Gets a comment from the collection based on its ID.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[assignTask(assignee: Identity)](/javascript/api/excel/excel.commentreply#assignTask_assignee_)|Assigns the task attached to the comment to the given user as the sole assignee.|
||[getTask()](/javascript/api/excel/excel.commentreply#getTask__)|Gets the task associated with this comment reply's thread.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.commentreply#getTaskOrNullObject__)|Gets the task associated with this comment reply's thread.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[getItemOrNullObject(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#getItemOrNullObject_commentReplyId_)|Returns a comment reply identified by its ID.|
|[ConditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|[getItemOrNullObject(id: string)](/javascript/api/excel/excel.conditionalformatcollection#getItemOrNullObject_id_)|Returns a conditional format identified by its ID.|
|[DocumentTask](/javascript/api/excel/excel.documenttask)|[percentComplete](/javascript/api/excel/excel.documenttask#percentComplete)|Specifies the completion percentage of the task.|
||[priority](/javascript/api/excel/excel.documenttask#priority)|Specifies the priority of the task.|
||[assignees](/javascript/api/excel/excel.documenttask#assignees)|Returns a collection of assignees of the task.|
||[changes](/javascript/api/excel/excel.documenttask#changes)|Gets the change records of the task.|
||[comment](/javascript/api/excel/excel.documenttask#comment)|Gets the comment associated with the task.|
||[completedBy](/javascript/api/excel/excel.documenttask#completedBy)|Gets the most recent user to have completed the task.|
||[completedDateTime](/javascript/api/excel/excel.documenttask#completedDateTime)|Gets the date and time that the task was completed.|
||[createdBy](/javascript/api/excel/excel.documenttask#createdBy)|Gets the user who created the task.|
||[createdDateTime](/javascript/api/excel/excel.documenttask#createdDateTime)|Gets the date and time that the task was created.|
||[id](/javascript/api/excel/excel.documenttask#id)|Gets the ID of the task.|
||[setStartAndDueDateTime(startDateTime: Date, dueDateTime: Date)](/javascript/api/excel/excel.documenttask#setStartAndDueDateTime_startDateTime__dueDateTime_)|Changes the start and the due dates of the task.|
||[startAndDueDateTime](/javascript/api/excel/excel.documenttask#startAndDueDateTime)|Gets or sets the date and time the task should start and is due.|
||[title](/javascript/api/excel/excel.documenttask#title)|Specifies title of the task.|
|[DocumentTaskChange](/javascript/api/excel/excel.documenttaskchange)|[assignee](/javascript/api/excel/excel.documenttaskchange#assignee)|Represents the user assigned to the task for an `assign` change record type, or the user unassigned from the task for an `unassign` change record type.|
||[changedBy](/javascript/api/excel/excel.documenttaskchange#changedBy)|Represents the user who created or changed the task.|
||[commentId](/javascript/api/excel/excel.documenttaskchange#commentId)|Represents the ID of the `Comment` or `CommentReply` to which the task change is anchored.|
||[createdDateTime](/javascript/api/excel/excel.documenttaskchange#createdDateTime)|Represents the creation date and time of the task change record.|
||[dueDateTime](/javascript/api/excel/excel.documenttaskchange#dueDateTime)|Represents the task's due date and time, in UTC time zone.|
||[id](/javascript/api/excel/excel.documenttaskchange#id)|ID for the task change record.|
||[percentComplete](/javascript/api/excel/excel.documenttaskchange#percentComplete)|Represents the task's completion percentage.|
||[priority](/javascript/api/excel/excel.documenttaskchange#priority)|Represents the task's priority.|
||[startDateTime](/javascript/api/excel/excel.documenttaskchange#startDateTime)|Represents the task's start date and time, in UTC time zone.|
||[title](/javascript/api/excel/excel.documenttaskchange#title)|Represents the task's title.|
||[type](/javascript/api/excel/excel.documenttaskchange#type)|Represents the action type of the task change record.|
||[undoHistoryId](/javascript/api/excel/excel.documenttaskchange#undoHistoryId)|Represents the `DocumentTaskChange.id` property that was undone for the `undo` change record type.|
|[DocumentTaskChangeCollection](/javascript/api/excel/excel.documenttaskchangecollection)|[getCount()](/javascript/api/excel/excel.documenttaskchangecollection#getCount__)|Gets the number of change records in the collection for the task.|
||[getItemAt(index: number)](/javascript/api/excel/excel.documenttaskchangecollection#getItemAt_index_)|Gets a task change record by using its index in the collection.|
||[items](/javascript/api/excel/excel.documenttaskchangecollection#items)|Gets the loaded child items in this collection.|
|[DocumentTaskCollection](/javascript/api/excel/excel.documenttaskcollection)|[getCount()](/javascript/api/excel/excel.documenttaskcollection#getCount__)|Gets the number of tasks in the collection.|
||[getItem(key: string)](/javascript/api/excel/excel.documenttaskcollection#getItem_key_)|Gets a task using its ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.documenttaskcollection#getItemAt_index_)|Gets a task by its index in the collection.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.documenttaskcollection#getItemOrNullObject_key_)|Gets a task using its ID.|
||[items](/javascript/api/excel/excel.documenttaskcollection#items)|Gets the loaded child items in this collection.|
|[DocumentTaskSchedule](/javascript/api/excel/excel.documenttaskschedule)|[dueDateTime](/javascript/api/excel/excel.documenttaskschedule#dueDateTime)|Gets the date and time that the task is due.|
||[startDateTime](/javascript/api/excel/excel.documenttaskschedule#startDateTime)|Gets the date and time that the task should start.|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.groupshapecollection#getItemOrNullObject_key_)|Gets a shape using its name or ID.|
|[Identity](/javascript/api/excel/excel.identity)|[displayName](/javascript/api/excel/excel.identity#displayName)|Represents the user's display name.|
||[email](/javascript/api/excel/excel.identity#email)|Represents the user's email address.|
||[id](/javascript/api/excel/excel.identity#id)|Represents the user's unique ID.|
|[IdentityCollection](/javascript/api/excel/excel.identitycollection)|[add(assignee: Identity)](/javascript/api/excel/excel.identitycollection#add_assignee_)|Adds a user identity to the collection.|
||[clear()](/javascript/api/excel/excel.identitycollection#clear__)|Removes all user identities from the collection.|
||[getCount()](/javascript/api/excel/excel.identitycollection#getCount__)|Gets the number of items in the collection.|
||[getItemAt(index: number)](/javascript/api/excel/excel.identitycollection#getItemAt_index_)|Gets a document user identity by using its index in the collection.|
||[items](/javascript/api/excel/excel.identitycollection#items)|Gets the loaded child items in this collection.|
||[remove(assignee: Identity)](/javascript/api/excel/excel.identitycollection#remove_assignee_)|Removes a user identity from the collection.|
|[IdentityEntity](/javascript/api/excel/excel.identityentity)|[displayName](/javascript/api/excel/excel.identityentity#displayName)|Represents the user's display name.|
||[email](/javascript/api/excel/excel.identityentity#email)|Represents the user's email address.|
||[id](/javascript/api/excel/excel.identityentity#id)|Represents the user's unique ID.|
|[LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|[dataProvider](/javascript/api/excel/excel.linkeddatatype#dataProvider)|The name of the data provider for the linked data type.|
||[lastRefreshed](/javascript/api/excel/excel.linkeddatatype#lastRefreshed)|The local time-zone date and time since the workbook was opened when the linked data type was last refreshed.|
||[name](/javascript/api/excel/excel.linkeddatatype#name)|The name of the linked data type.|
||[periodicRefreshInterval](/javascript/api/excel/excel.linkeddatatype#periodicRefreshInterval)|The frequency, in seconds, at which the linked data type is refreshed if `refreshMode` is set to "Periodic".|
||[refreshMode](/javascript/api/excel/excel.linkeddatatype#refreshMode)|The mechanism by which the data for the linked data type is retrieved.|
||[serviceId](/javascript/api/excel/excel.linkeddatatype#serviceId)|The unique ID of the linked data type.|
||[supportedRefreshModes](/javascript/api/excel/excel.linkeddatatype#supportedRefreshModes)|Returns an array with all the refresh modes supported by the linked data type.|
||[requestRefresh()](/javascript/api/excel/excel.linkeddatatype#requestRefresh__)|Makes a request to refresh the linked data type.|
||[requestSetRefreshMode(refreshMode: Excel.LinkedDataTypeRefreshMode)](/javascript/api/excel/excel.linkeddatatype#requestSetRefreshMode_refreshMode_)|Makes a request to change the refresh mode for this linked data type.|
|[LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs)|[serviceId](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#serviceId)|The unique ID of the new linked data type.|
||[source](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#source)|Gets the source of the event.|
||[type](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#type)|Gets the type of the event.|
|[LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection)|[getCount()](/javascript/api/excel/excel.linkeddatatypecollection#getCount__)|Gets the number of linked data types in the collection.|
||[getItem(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getItem_key_)|Gets a linked data type by service ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.linkeddatatypecollection#getItemAt_index_)|Gets a linked data type by its index in the collection.|
||[getItemOrNullObject(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getItemOrNullObject_key_)|Gets a linked data type by ID.|
||[items](/javascript/api/excel/excel.linkeddatatypecollection#items)|Gets the loaded child items in this collection.|
||[requestRefreshAll()](/javascript/api/excel/excel.linkeddatatypecollection#requestRefreshAll__)|Makes a request to refresh all the linked data types in the collection.|
|[LinkedWorkbook](/javascript/api/excel/excel.linkedworkbook)|[breakLinks()](/javascript/api/excel/excel.linkedworkbook#breakLinks__)|Makes a request to break the links pointing to the linked workbook.|
||[id](/javascript/api/excel/excel.linkedworkbook#id)|The original URL pointing to the linked workbook.|
||[refresh()](/javascript/api/excel/excel.linkedworkbook#refresh__)|Makes a request to refresh the data retrieved from the linked workbook.|
|[LinkedWorkbookCollection](/javascript/api/excel/excel.linkedworkbookcollection)|[breakAllLinks()](/javascript/api/excel/excel.linkedworkbookcollection#breakAllLinks__)|Breaks all the links to the linked workbooks.|
||[getItem(key: string)](/javascript/api/excel/excel.linkedworkbookcollection#getItem_key_)|Gets information about a linked workbook by its URL.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.linkedworkbookcollection#getItemOrNullObject_key_)|Gets information about a linked workbook by its URL.|
||[items](/javascript/api/excel/excel.linkedworkbookcollection#items)|Gets the loaded child items in this collection.|
||[refreshAll()](/javascript/api/excel/excel.linkedworkbookcollection#refreshAll__)|Makes a request to refresh all the workbook links.|
||[workbookLinksRefreshMode](/javascript/api/excel/excel.linkedworkbookcollection#workbookLinksRefreshMode)|Represents the update mode of the workbook links.|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getItemOrNullObject_key_)|Gets a sheet view using its name.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getCell_dataHierarchy__rowItems__columnItems_)|Gets a unique cell in the PivotTable based on a data hierarchy and the row and column items of their respective hierarchies.|
||[pivotStyle](/javascript/api/excel/excel.pivotlayout#pivotStyle)|The style applied to the PivotTable.|
||[setStyle(style: string \| PivotTableStyle \| BuiltInPivotTableStyle)](/javascript/api/excel/excel.pivotlayout#setStyle_style_)|Sets the style applied to the PivotTable.|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getFirstOrNullObject()](/javascript/api/excel/excel.pivottablescopedcollection#getFirstOrNullObject__)|Gets the first PivotTable in the collection.|
|[Query](/javascript/api/excel/excel.query)|[error](/javascript/api/excel/excel.query#error)|Gets the query error message from when the query was last refreshed.|
||[loadedTo](/javascript/api/excel/excel.query#loadedTo)|Gets the query 'loaded to' object type.|
||[loadedToDataModel](/javascript/api/excel/excel.query#loadedToDataModel)|Specifies if the query loaded to the Data Model.|
||[name](/javascript/api/excel/excel.query#name)|Gets the name of the query.|
||[refreshDate](/javascript/api/excel/excel.query#refreshDate)|Gets the date and time when the query was last refreshed.|
||[rowsLoadedCount](/javascript/api/excel/excel.query#rowsLoadedCount)|Gets the number of rows that were loaded when the query was last refreshed.|
|[QueryCollection](/javascript/api/excel/excel.querycollection)|[getCount()](/javascript/api/excel/excel.querycollection#getCount__)|Gets the number of queries in the workbook.|
||[getItem(key: string)](/javascript/api/excel/excel.querycollection#getItem_key_)|Gets a query from the collection based on its name.|
||[items](/javascript/api/excel/excel.querycollection#items)|Gets the loaded child items in this collection.|
|[Range](/javascript/api/excel/excel.range)|[getDependents()](/javascript/api/excel/excel.range#getDependents__)|Returns a `WorkbookRangeAreas` object that represents the range containing all the dependents of a cell in the same worksheet or in multiple worksheets.|
||[getPrecedents()](/javascript/api/excel/excel.range#getPrecedents__)|Returns a `WorkbookRangeAreas` object that represents the range containing all the precedents of a cell in the same worksheet or in multiple worksheets.|
|[RefreshModeChangedEventArgs](/javascript/api/excel/excel.refreshmodechangedeventargs)|[refreshMode](/javascript/api/excel/excel.refreshmodechangedeventargs#refreshMode)|The linked data type refresh mode.|
||[serviceId](/javascript/api/excel/excel.refreshmodechangedeventargs#serviceId)|The unique ID of the object whose refresh mode was changed.|
||[source](/javascript/api/excel/excel.refreshmodechangedeventargs#source)|Gets the source of the event.|
||[type](/javascript/api/excel/excel.refreshmodechangedeventargs#type)|Gets the type of the event.|
|[RefreshRequestCompletedEventArgs](/javascript/api/excel/excel.refreshrequestcompletedeventargs)|[refreshed](/javascript/api/excel/excel.refreshrequestcompletedeventargs#refreshed)|Indicates if the request to refresh was successful.|
||[serviceId](/javascript/api/excel/excel.refreshrequestcompletedeventargs#serviceId)|The unique ID of the object whose refresh request was completed.|
||[source](/javascript/api/excel/excel.refreshrequestcompletedeventargs#source)|Gets the source of the event.|
||[type](/javascript/api/excel/excel.refreshrequestcompletedeventargs#type)|Gets the type of the event.|
||[warnings](/javascript/api/excel/excel.refreshrequestcompletedeventargs#warnings)|An array that contains any warnings generated from the refresh request.|
|[Shape](/javascript/api/excel/excel.shape)|[displayName](/javascript/api/excel/excel.shape#displayName)|Gets the display name of the shape.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addSvg_xml_)|Creates a scalable vector graphic (SVG) from an XML string and adds it to the worksheet.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.shapecollection#getItemOrNullObject_key_)|Gets a shape using its name or ID.|
|[Slicer](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameInFormula)|Represents the slicer name used in the formula.|
||[slicerStyle](/javascript/api/excel/excel.slicer#slicerStyle)|The style applied to the slicer.|
||[setStyle(style: string \| SlicerStyle \| BuiltInSlicerStyle)](/javascript/api/excel/excel.slicer#setStyle_style_)|Sets the style applied to the slicer.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getItemOrNullObject(name: string)](/javascript/api/excel/excel.stylecollection#getItemOrNullObject_name_)|Gets a style by name.|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearStyle__)|Changes the table to use the default table style.|
||[onFiltered](/javascript/api/excel/excel.table#onFiltered)|Occurs when a filter is applied on a specific table.|
||[tableStyle](/javascript/api/excel/excel.table#tableStyle)|The style applied to the table.|
||[setStyle(style: string \| TableStyle \| BuiltInTableStyle)](/javascript/api/excel/excel.table#setStyle_style_)|Sets the style applied to the table.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onFiltered)|Occurs when a filter is applied on any table in a workbook, or a worksheet.|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableId)|Gets the ID of the table in which the filter is applied.|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetId)|Gets the ID of the worksheet which contains the table.|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[deleteRows(rows: number[] \| TableRow[])](/javascript/api/excel/excel.tablerowcollection#deleteRows_rows_)|Delete multiple rows from a table.|
||[deleteRowsAt(index: number, count?: number)](/javascript/api/excel/excel.tablerowcollection#deleteRowsAt_index__count_)|Delete a specified number of rows from a table, starting at a given index.|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.tablescopedcollection#getItemOrNullObject_key_)|Gets a table by name or ID.|
|[Workbook](/javascript/api/excel/excel.workbook)|[linkedDataTypes](/javascript/api/excel/excel.workbook#linkedDataTypes)|Returns a collection of linked data types that are part of the workbook.|
||[linkedWorkbooks](/javascript/api/excel/excel.workbook#linkedWorkbooks)|Returns a collection of linked workbooks.|
||[queries](/javascript/api/excel/excel.workbook#queries)|Returns a collection of Power Query queries that are part of the workbook.|
||[tasks](/javascript/api/excel/excel.workbook#tasks)|Returns a collection of tasks that are present in the workbook.|
||[showPivotFieldList](/javascript/api/excel/excel.workbook#showPivotFieldList)|Specifies whether the PivotTable's field list pane is shown at the workbook level.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904DateSystem)|True if the workbook uses the 1904 date system.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onFiltered](/javascript/api/excel/excel.worksheet#onFiltered)|Occurs when a filter is applied on a specific worksheet.|
||[onProtectionChanged](/javascript/api/excel/excel.worksheet#onProtectionChanged)|Occurs when the worksheet protection state is changed.|
||[tabId](/javascript/api/excel/excel.worksheet#tabId)|Returns a value representing this worksheet that can be read by Open Office XML.|
||[tasks](/javascript/api/excel/excel.worksheet#tasks)|Returns a collection of tasks that are present in the worksheet.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[changeDirectionState](/javascript/api/excel/excel.worksheetchangedeventargs#changeDirectionState)|Represents a change to the direction that the cells in a worksheet will shift when a cell or cells are deleted or inserted.|
||[triggerSource](/javascript/api/excel/excel.worksheetchangedeventargs#triggerSource)|Represents the trigger source of the event.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addFromBase64_base64File__sheetNamesToInsert__positionType__relativeTo_)|Inserts the specified worksheets of a workbook into the current workbook.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onFiltered)|Occurs when any worksheet's filter is applied in the workbook.|
||[onProtectionChanged](/javascript/api/excel/excel.worksheetcollection#onProtectionChanged)|Occurs when the worksheet protection state is changed.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetId)|Gets the ID of the worksheet in which the filter is applied.|
|[WorksheetProtectionChangedEventArgs](/javascript/api/excel/excel.worksheetprotectionchangedeventargs)|[isProtected](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#isProtected)|Gets the current protection status of the worksheet.|
||[source](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#source)|The source of the event.|
||[type](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#type)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#worksheetId)|Gets the ID of the worksheet in which the protection status is changed.|

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel?view=excel-js-preview&preserve-view=true)
- [Excel JavaScript API requirement sets](excel-api-requirement-sets.md)
