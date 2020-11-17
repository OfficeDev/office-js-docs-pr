---
title: Excel JavaScript preview APIs
description: 'Details about upcoming Excel JavaScript APIs.'
ms.date: 11/17/2020
ms.prod: excel
localization_priority: Normal
---

# Excel JavaScript preview APIs

New Excel JavaScript APIs are first introduced in "preview" and later become part of a specific, numbered requirement set after sufficient testing occurs and user feedback is acquired.

The first table provides a concise summary of the APIs, while the subsequent table gives a detailed list.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| Feature area | Description | Relevant objects |
|:--- |:--- |:--- |
| Linked data types | Adds support for data types connected to Excel from external sources. | [LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|
| Named sheet views | Gives programmatic control of per-user worksheet views. | [NamedSheetView](/javascript/api/excel/excel.namedsheetview) |
| Tasks | Turn comments into tasks assigned to users. | [Task](/javascript/api/excel/excel.task) |

## API list

The following table lists the Excel JavaScript APIs currently in preview. For a complete list of all Excel JavaScript APIs (including preview APIs and previously released APIs), see [all Excel JavaScript APIs](/javascript/api/excel?view=excel-js-preview&preserve-view=true).

| Class | Fields | Description |
|:---|:---|:---|
|[Comment](/javascript/api/excel/excel.comment)|[assignTask(email: string)](/javascript/api/excel/excel.comment#assigntask-email-)|Assigns the task attached to the comment to the given user as the sole assignee.|
||[getTask()](/javascript/api/excel/excel.comment#gettask--)|Gets the task associated with this comment.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.comment#gettaskornullobject--)|Gets the task associated with this comment.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[assignTask(email: string)](/javascript/api/excel/excel.commentreply#assigntask-email-)|Assigns the task attached to the comment to the given user as the sole assignee.|
||[getTask()](/javascript/api/excel/excel.commentreply#gettask--)|Gets the task associated with this comment.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.commentreply#gettaskornullobject--)|Gets the task associated with this comment.|
|[LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|[dataProvider](/javascript/api/excel/excel.linkeddatatype#dataprovider)|The name of the data provider for the linked data type.|
||[lastRefreshed](/javascript/api/excel/excel.linkeddatatype#lastrefreshed)|The local time-zone date and time since the workbook was opened when the linked data type was last refreshed.|
||[name](/javascript/api/excel/excel.linkeddatatype#name)|The name of the linked data type.|
||[periodicRefreshInterval](/javascript/api/excel/excel.linkeddatatype#periodicrefreshinterval)|The frequency, in seconds, at which the linked data type is refreshed if `refreshMode` is set to "Periodic".|
||[refreshMode](/javascript/api/excel/excel.linkeddatatype#refreshmode)|The mechanism by which the data for the linked data type is retrieved.|
||[serviceId](/javascript/api/excel/excel.linkeddatatype#serviceid)|The unique id of the linked data type.|
||[supportedRefreshModes](/javascript/api/excel/excel.linkeddatatype#supportedrefreshmodes)|Returns an array with all the refresh modes supported by the linked data type.|
||[requestRefresh()](/javascript/api/excel/excel.linkeddatatype#requestrefresh--)|Makes a request to refresh the linked data type.|
||[requestSetRefreshMode(refreshMode: Excel.LinkedDataTypeRefreshMode)](/javascript/api/excel/excel.linkeddatatype#requestsetrefreshmode-refreshmode-)|Makes a request to change the refresh mode for this linked data type.|
|[LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs)|[serviceId](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#serviceid)|The unique id of the new linked data type.|
||[source](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#source)|Gets the source of the event.|
||[type](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#type)|Gets the type of the event.|
|[LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection)|[getCount()](/javascript/api/excel/excel.linkeddatatypecollection#getcount--)|Gets the number of linked data types in the collection.|
||[getItem(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitem-key-)|Gets a linked data type by service id.|
||[getItemAt(index: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitemat-index-)|Gets a linked data type by its index in the collection.|
||[getItemOrNullObject(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitemornullobject-key-)|Gets a linked data type by ID.|
||[items](/javascript/api/excel/excel.linkeddatatypecollection#items)|Gets the loaded child items in this collection.|
||[requestRefreshAll()](/javascript/api/excel/excel.linkeddatatypecollection#requestrefreshall--)|Makes a request to refresh all the linked data types in the collection.|
|[NamedSheetView](/javascript/api/excel/excel.namedsheetview)|[activate()](/javascript/api/excel/excel.namedsheetview#activate--)|Activates this sheet view.|
||[delete()](/javascript/api/excel/excel.namedsheetview#delete--)|Removes the sheet view from the worksheet.|
||[duplicate(name?: string)](/javascript/api/excel/excel.namedsheetview#duplicate-name-)|Creates a copy of this sheet view.|
||[name](/javascript/api/excel/excel.namedsheetview#name)|Gets or sets the name of the sheet view.|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[add(name: string)](/javascript/api/excel/excel.namedsheetviewcollection#add-name-)|Creates a new sheet view with the given name.|
||[enterTemporary()](/javascript/api/excel/excel.namedsheetviewcollection#entertemporary--)|Creates and activates a new temporary sheet view.|
||[exit()](/javascript/api/excel/excel.namedsheetviewcollection#exit--)|Exits the currently active sheet view.|
||[getActive()](/javascript/api/excel/excel.namedsheetviewcollection#getactive--)|Gets the worksheet's currently active sheet view.|
||[getCount()](/javascript/api/excel/excel.namedsheetviewcollection#getcount--)|Gets the number of sheet views in this worksheet.|
||[getItem(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getitem-key-)|Gets a sheet view using its name.|
||[getItemAt(index: number)](/javascript/api/excel/excel.namedsheetviewcollection#getitemat-index-)|Gets a sheet view by its index in the collection.|
||[items](/javascript/api/excel/excel.namedsheetviewcollection#items)|Gets the loaded child items in this collection.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[altTextDescription](/javascript/api/excel/excel.pivotlayout#alttextdescription)|The alt text description of the PivotTable.|
||[altTextTitle](/javascript/api/excel/excel.pivotlayout#alttexttitle)|The alt text title of the PivotTable.|
||[displayBlankLineAfterEachItem(display: boolean)](/javascript/api/excel/excel.pivotlayout#displayblanklineaftereachitem-display-)|Sets whether or not to display a blank line after each item.|
||[emptyCellText](/javascript/api/excel/excel.pivotlayout#emptycelltext)|The text that is automatically filled into any empty cell in the PivotTable if `fillEmptyCells == true`.|
||[fillEmptyCells](/javascript/api/excel/excel.pivotlayout#fillemptycells)|Specifies whether empty cells in the PivotTable should be populated with the `emptyCellText`.|
||[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|Gets a unique cell in the PivotTable based on a data hierarchy and the row and column items of their respective hierarchies.|
||[pivotStyle](/javascript/api/excel/excel.pivotlayout#pivotstyle)|The style applied to the PivotTable.|
||[repeatAllItemLabels(repeatLabels: boolean)](/javascript/api/excel/excel.pivotlayout#repeatallitemlabels-repeatlabels-)|Sets the "repeat all item labels" setting across all fields in the PivotTable.|
||[setStyle(style: string \| PivotTableStyle \| BuiltInPivotTableStyle)](/javascript/api/excel/excel.pivotlayout#setstyle-style-)|Sets the style applied to the PivotTable.|
||[showFieldHeaders](/javascript/api/excel/excel.pivotlayout#showfieldheaders)|Specifies whether the PivotTable displays field headers (field captions and filter drop-downs).|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[refreshOnOpen](/javascript/api/excel/excel.pivottable#refreshonopen)|Specifies whether the PivotTable refreshes when the workbook opens.|
|[Range](/javascript/api/excel/excel.range)|[getPrecedents()](/javascript/api/excel/excel.range#getprecedents--)|Returns a `WorkbookRangeAreas` object that represents the range containing all the precedents of a cell in same worksheet or in multiple worksheets.|
|[RefreshModeChangedEventArgs](/javascript/api/excel/excel.refreshmodechangedeventargs)|[refreshMode](/javascript/api/excel/excel.refreshmodechangedeventargs#refreshmode)|The linked data type refresh mode.|
||[serviceId](/javascript/api/excel/excel.refreshmodechangedeventargs#serviceid)|The unique id of the object whose refresh mode was changed.|
||[source](/javascript/api/excel/excel.refreshmodechangedeventargs#source)|Gets the source of the event.|
||[type](/javascript/api/excel/excel.refreshmodechangedeventargs#type)|Gets the type of the event.|
|[RefreshRequestCompletedEventArgs](/javascript/api/excel/excel.refreshrequestcompletedeventargs)|[refreshed](/javascript/api/excel/excel.refreshrequestcompletedeventargs#refreshed)|Indicates if the request to refresh was successful.|
||[serviceId](/javascript/api/excel/excel.refreshrequestcompletedeventargs#serviceid)|The unique id of the object whose refresh request was completed.|
||[source](/javascript/api/excel/excel.refreshrequestcompletedeventargs#source)|Gets the source of the event.|
||[type](/javascript/api/excel/excel.refreshrequestcompletedeventargs#type)|Gets the type of the event.|
||[warnings](/javascript/api/excel/excel.refreshrequestcompletedeventargs#warnings)|An array that contains any warnings generated from the refresh request.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|Creates a scalable vector graphic (SVG) from an XML string and adds it to the worksheet.|
|[Slicer](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|Represents the slicer name used in the formula.|
||[slicerStyle](/javascript/api/excel/excel.slicer#slicerstyle)|The style applied to the Slicer.|
||[setStyle(style: string \| SlicerStyle \| BuiltInSlicerStyle)](/javascript/api/excel/excel.slicer#setstyle-style-)|Sets the style applied to the slicer.|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|Changes the table to use the default table style.|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|Occurs when filter is applied on a specific table.|
||[tableStyle](/javascript/api/excel/excel.table#tablestyle)|The style applied to the Table.|
||[setStyle(style: string \| TableStyle \| BuiltInTableStyle)](/javascript/api/excel/excel.table#setstyle-style-)|Sets the style applied to the table.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|Occurs when filter is applied on any table in a workbook, or a worksheet.|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|Gets the id of the table in which the filter is applied.|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|Gets the id of the worksheet which contains the table.|
|[Task](/javascript/api/excel/excel.task)|[addAssignee(email: string)](/javascript/api/excel/excel.task#addassignee-email-)|Adds an assignee to the task.|
||[applyChanges(taskChanges: Excel.TaskChanges)](/javascript/api/excel/excel.task#applychanges-taskchanges-)|Applies the given changes to the task.|
||[assignees](/javascript/api/excel/excel.task#assignees)|Gets the users to whom the task is assigned.|
||[comment](/javascript/api/excel/excel.task#comment)|Gets the comment associated with the task.|
||[dueDate](/javascript/api/excel/excel.task#duedate)|Gets the date and time the task is due.|
||[historyRecords](/javascript/api/excel/excel.task#historyrecords)|Gets the history records of the task.|
||[id](/javascript/api/excel/excel.task#id)|Gets the id of the task.|
||[percentComplete](/javascript/api/excel/excel.task#percentcomplete)|Gets the completion percentage of the task.|
||[priority](/javascript/api/excel/excel.task#priority)|Gets the priority of the task.|
||[startDate](/javascript/api/excel/excel.task#startdate)|Gets the date and time the task should start.|
||[title](/javascript/api/excel/excel.task#title)|Gets title of the task.|
||[removeAllAssignees()](/javascript/api/excel/excel.task#removeallassignees--)|Removes all assignees from the task.|
||[removeAssignee(email: string)](/javascript/api/excel/excel.task#removeassignee-email-)|Removes an assignee from the task.|
||[setPercentComplete(percentComplete: number)](/javascript/api/excel/excel.task#setpercentcomplete-percentcomplete-)|Changes the completion of the task.|
||[setPriority(priority: number)](/javascript/api/excel/excel.task#setpriority-priority-)|Changes the priority of the task.|
||[setStartDateAndDueDate(startDate: Date, dueDate: Date)](/javascript/api/excel/excel.task#setstartdateandduedate-startdate--duedate-)|Changes the start and the due dates of the task.|
||[setTitle(title: string)](/javascript/api/excel/excel.task#settitle-title-)|Changes the title of the task.|
|[TaskChanges](/javascript/api/excel/excel.taskchanges)|[dueDate](/javascript/api/excel/excel.taskchanges#duedate)|Sets a new due date for the task, in UTC time zone.|
||[emailsToAssign](/javascript/api/excel/excel.taskchanges#emailstoassign)|Sets email addresses of the users to assign to the task.|
||[emailsToUnassign](/javascript/api/excel/excel.taskchanges#emailstounassign)|Sets email addresses of the users to unassign from the task.|
||[percentComplete](/javascript/api/excel/excel.taskchanges#percentcomplete)|Sets a new completion percentage for the task.|
||[priority](/javascript/api/excel/excel.taskchanges#priority)|Sets a new priority for the task.|
||[removeAllPreviousAssignees](/javascript/api/excel/excel.taskchanges#removeallpreviousassignees)|Sets if the change should remove all previous assignees from the task.|
||[startDate](/javascript/api/excel/excel.taskchanges#startdate)|Sets a new start date for the task, in UTC time zone.|
||[title](/javascript/api/excel/excel.taskchanges#title)|Sets a new title for the task.|
|[TaskCollection](/javascript/api/excel/excel.taskcollection)|[getCount()](/javascript/api/excel/excel.taskcollection#getcount--)|Gets the number of tasks in the collection.|
||[getItem(key: string)](/javascript/api/excel/excel.taskcollection#getitem-key-)|Gets a task using its id.|
||[getItemAt(index: number)](/javascript/api/excel/excel.taskcollection#getitemat-index-)|Gets a task by its index in the collection.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.taskcollection#getitemornullobject-key-)|Gets a task using its id.|
||[items](/javascript/api/excel/excel.taskcollection#items)|Gets the loaded child items in this collection.|
|[TaskHistoryRecord](/javascript/api/excel/excel.taskhistoryrecord)|[anchorId](/javascript/api/excel/excel.taskhistoryrecord#anchorid)|Represents the ID of the object to which the task is anchored (e.g., commentId for tasks attached to comments).|
||[assignee](/javascript/api/excel/excel.taskhistoryrecord#assignee)|Represents the user assigned to the task for an "Assign" history record type, or the user to unassign from the task for an "Unassign" history record type.|
||[attributionUser](/javascript/api/excel/excel.taskhistoryrecord#attributionuser)|Represents the user who created or changed the task.|
||[dueDate](/javascript/api/excel/excel.taskhistoryrecord#duedate)|Represents the task's due date.|
||[historyRecordCreatedDate](/javascript/api/excel/excel.taskhistoryrecord#historyrecordcreateddate)|Represents creation date of the task history record.|
||[id](/javascript/api/excel/excel.taskhistoryrecord#id)|ID for the history record.|
||[percentComplete](/javascript/api/excel/excel.taskhistoryrecord#percentcomplete)|Represents the task's completion percentage.|
||[priority](/javascript/api/excel/excel.taskhistoryrecord#priority)|Represents the task's priority.|
||[startDate](/javascript/api/excel/excel.taskhistoryrecord#startdate)|Represents the task's start date.|
||[title](/javascript/api/excel/excel.taskhistoryrecord#title)|Represents the task's title.|
||[type](/javascript/api/excel/excel.taskhistoryrecord#type)|Represents task history record's type.|
||[undoHistoryId](/javascript/api/excel/excel.taskhistoryrecord#undohistoryid)|Represents the TaskHistoryRecord.id property that was undone for the "Undo" history record type.|
|[TaskHistoryRecordCollection](/javascript/api/excel/excel.taskhistoryrecordcollection)|[getCount()](/javascript/api/excel/excel.taskhistoryrecordcollection#getcount--)|Gets the number of history records in the collection for the task.|
||[getItemAt(index: number)](/javascript/api/excel/excel.taskhistoryrecordcollection#getitemat-index-)|Gets a task history record by using its index in the collection.|
||[items](/javascript/api/excel/excel.taskhistoryrecordcollection#items)|Gets the loaded child items in this collection.|
|[User](/javascript/api/excel/excel.user)|[displayName](/javascript/api/excel/excel.user#displayname)|Represents the user's display name.|
||[email](/javascript/api/excel/excel.user#email)|Represents the user's email address.|
||[uid](/javascript/api/excel/excel.user#uid)|Represents the user's unique ID.|
|[Workbook](/javascript/api/excel/excel.workbook)|[linkedDataTypes](/javascript/api/excel/excel.workbook#linkeddatatypes)|Returns a collection of linked data types that are part of the workbook.|
||[tasks](/javascript/api/excel/excel.workbook#tasks)|Returns a collection of tasks that are present in the workbook.|
||[showPivotFieldList](/javascript/api/excel/excel.workbook#showpivotfieldlist)|Specifies whether the PivotTable's field list pane is shown at the workbook level.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|True if the workbook uses the 1904 date system.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[namedSheetViews](/javascript/api/excel/excel.worksheet#namedsheetviews)|Returns a collection of sheet views that are present in the worksheet.|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|Occurs when filter is applied on a specific worksheet.|
||[tasks](/javascript/api/excel/excel.worksheet#tasks)|Returns a collection of tasks that are present in the worksheet.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|Inserts the specified worksheets of a workbook into the current workbook.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|Occurs when any worksheet's filter is applied in the workbook.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|Gets the id of the worksheet in which the filter is applied.|

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel?view=excel-js-preview&preserve-view=true)
- [Excel JavaScript API requirement sets](excel-api-requirement-sets.md)
