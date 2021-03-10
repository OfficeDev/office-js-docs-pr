---
title: Excel JavaScript preview APIs
description: 'Details about upcoming Excel JavaScript APIs.'
ms.date: 03/10/2021
ms.prod: excel
localization_priority: Normal
---

# Excel JavaScript preview APIs

New Excel JavaScript APIs are first introduced in "preview" and later become part of a specific, numbered requirement set after sufficient testing occurs and user feedback is acquired.

The first table provides a concise summary of the APIs, while the subsequent table gives a detailed list.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| Feature area | Description | Relevant objects |
|:--- |:--- |:--- |
| Document tasks | Turn comments into tasks assigned to users. | [DocumentTask](/javascript/api/excel/excel.documenttask) |
| Formula changed events | Track changes to formulas, including the source and type of event that caused a change. | [Worksheet.onFormulaChanged](/javascript/api/excel/excel.worksheet#onFormulaChanged)|
| Linked data types | Adds support for data types connected to Excel from external sources. | [LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|
| Named sheet views | Gives programmatic control of per-user worksheet views. | [NamedSheetView](/javascript/api/excel/excel.namedsheetview) |
| PivotTable PivotLayout | Expansion of the PivotLayout class, including new support for alt text and empty cell management. | [PivotLayout](/javascript/api/excel/excel.pivotlayout) |
| Table styles | Control font, border, fill color, and other table styles. | [Table](/javascript/api/excel/excel.table), [PivotTable](/javascript/api/excel/excel.pivottable), [Slicer](/javascript/api/excel/excel.slicer) |

## API list

The following table lists the Excel JavaScript APIs currently in preview. For a complete list of all Excel JavaScript APIs (including preview APIs and previously released APIs), see [all Excel JavaScript APIs](/javascript/api/excel?view=excel-js-preview&preserve-view=true).

| Class | Fields | Description |
|:---|:---|:---|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[clearColumnCriteria(columnIndex: number)](/javascript/api/excel/excel.autofilter#clearcolumncriteria-columnindex-)|Clears the filter criteria of the AutoFilter.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)||[ChartCollectionCustom](/javascript/api/excel/excel.chartcollectioncustom)||[ChartFill](/javascript/api/excel/excel.chartfill)||[ChartFillCustom](/javascript/api/excel/excel.chartfillcustom)||[Comment](/javascript/api/excel/excel.comment)|[assignTask(assignee: Identity)](/javascript/api/excel/excel.comment#assigntask-assignee-)|Assigns the task attached to the comment to the given user as an assignee.|
||[getTask()](/javascript/api/excel/excel.comment#gettask--)|Gets the task associated with this comment.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.comment#gettaskornullobject--)|Gets the task associated with this comment.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[getItemOrNullObject(commentId: string)](/javascript/api/excel/excel.commentcollection#getitemornullobject-commentid-)|Gets a comment from the collection based on its ID.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)||[CommentCollectionCustom](/javascript/api/excel/excel.commentcollectioncustom)||[CommentReply](/javascript/api/excel/excel.commentreply)|[assignTask(assignee: Identity)](/javascript/api/excel/excel.commentreply#assigntask-assignee-)|Assigns the task attached to the comment to the given user as the sole assignee.|
||[getTask()](/javascript/api/excel/excel.commentreply#gettask--)|Gets the task associated with this comment reply's thread.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.commentreply#gettaskornullobject--)|Gets the task associated with this comment reply's thread.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[getItemOrNullObject(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#getitemornullobject-commentreplyid-)|Returns a comment reply identified by its ID.|
|[ConditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|[getItemOrNullObject(id: string)](/javascript/api/excel/excel.conditionalformatcollection#getitemornullobject-id-)|Returns a conditional format identified by its ID.|
|[DocumentTask](/javascript/api/excel/excel.documenttask)|[percentComplete](/javascript/api/excel/excel.documenttask#percentcomplete)|Specifies the completion percentage of the task.|
||[priority](/javascript/api/excel/excel.documenttask#priority)|Specifies the priority of the task.|
||[assignees](/javascript/api/excel/excel.documenttask#assignees)|Returns a collection of assignees of the task.|
||[changes](/javascript/api/excel/excel.documenttask#changes)|Gets the change records of the task.|
||[comment](/javascript/api/excel/excel.documenttask#comment)|Gets the comment associated with the task.|
||[completedBy](/javascript/api/excel/excel.documenttask#completedby)|Gets the most recent user to have completed the task.|
||[completedDateTime](/javascript/api/excel/excel.documenttask#completeddatetime)|Gets the date and time that the task was completed.|
||[createdBy](/javascript/api/excel/excel.documenttask#createdby)|Gets the user who created the task.|
||[createdDateTime](/javascript/api/excel/excel.documenttask#createddatetime)|Gets the date and time that the task was created.|
||[id](/javascript/api/excel/excel.documenttask#id)|Gets the ID of the task.|
||[setStartAndDueDateTime(startDateTime: Date, dueDateTime: Date)](/javascript/api/excel/excel.documenttask#setstartandduedatetime-startdatetime--duedatetime-)|Changes the start and the due dates of the task.|
||[startAndDueDateTime](/javascript/api/excel/excel.documenttask#startandduedatetime)|Gets or sets the date and time the task should start and is due.|
||[title](/javascript/api/excel/excel.documenttask#title)|Specifies title of the task.|
|[DocumentTaskChange](/javascript/api/excel/excel.documenttaskchange)|[assignee](/javascript/api/excel/excel.documenttaskchange#assignee)|Represents the user assigned to the task for an `assign` change record type, or the user unassigned from the task for an `unassign` change record type.|
||[changedBy](/javascript/api/excel/excel.documenttaskchange#changedby)|Represents the user who created or changed the task.|
||[commentId](/javascript/api/excel/excel.documenttaskchange#commentid)|Represents the ID of the `Comment` or `CommentReply` to which the task change is anchored.|
||[createdDateTime](/javascript/api/excel/excel.documenttaskchange#createddatetime)|Represents the creation date and time of the task change record.|
||[dueDateTime](/javascript/api/excel/excel.documenttaskchange#duedatetime)|Represents the task's due date and time, in UTC time zone.|
||[id](/javascript/api/excel/excel.documenttaskchange#id)|ID for the task change record.|
||[percentComplete](/javascript/api/excel/excel.documenttaskchange#percentcomplete)|Represents the task's completion percentage.|
||[priority](/javascript/api/excel/excel.documenttaskchange#priority)|Represents the task's priority.|
||[startDateTime](/javascript/api/excel/excel.documenttaskchange#startdatetime)|Represents the task's start date and time, in UTC time zone.|
||[title](/javascript/api/excel/excel.documenttaskchange#title)|Represents the task's title.|
||[type](/javascript/api/excel/excel.documenttaskchange#type)|Represents the action type of the task change record.|
||[undoHistoryId](/javascript/api/excel/excel.documenttaskchange#undohistoryid)|Represents the `DocumentTaskChange.id` property that was undone for the `undo` change record type.|
|[DocumentTaskChangeCollection](/javascript/api/excel/excel.documenttaskchangecollection)|[getCount()](/javascript/api/excel/excel.documenttaskchangecollection#getcount--)|Gets the number of change records in the collection for the task.|
||[getItemAt(index: number)](/javascript/api/excel/excel.documenttaskchangecollection#getitemat-index-)|Gets a task change record by using its index in the collection.|
||[items](/javascript/api/excel/excel.documenttaskchangecollection#items)|Gets the loaded child items in this collection.|
|[DocumentTaskCollection](/javascript/api/excel/excel.documenttaskcollection)|[getCount()](/javascript/api/excel/excel.documenttaskcollection#getcount--)|Gets the number of tasks in the collection.|
||[getItem(key: string)](/javascript/api/excel/excel.documenttaskcollection#getitem-key-)|Gets a task using its ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.documenttaskcollection#getitemat-index-)|Gets a task by its index in the collection.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.documenttaskcollection#getitemornullobject-key-)|Gets a task using its ID.|
||[items](/javascript/api/excel/excel.documenttaskcollection#items)|Gets the loaded child items in this collection.|
|[DocumentTaskSchedule](/javascript/api/excel/excel.documenttaskschedule)|[dueDateTime](/javascript/api/excel/excel.documenttaskschedule#duedatetime)|Gets the date and time that the task is due.|
||[startDateTime](/javascript/api/excel/excel.documenttaskschedule#startdatetime)|Gets the date and time that the task should start.|
|[FormulaChangedEventDetail](/javascript/api/excel/excel.formulachangedeventdetail)|[cellAddress](/javascript/api/excel/excel.formulachangedeventdetail#celladdress)|The address of the cell that contains the changed formula.|
||[previousFormula](/javascript/api/excel/excel.formulachangedeventdetail#previousformula)|Represents the previous formula, before it was changed.|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.groupshapecollection#getitemornullobject-key-)|Gets a shape using its name or ID.|
|[Identity](/javascript/api/excel/excel.identity)|[displayName](/javascript/api/excel/excel.identity#displayname)|Represents the user's display name.|
||[email](/javascript/api/excel/excel.identity#email)|Represents the user's email address.|
||[id](/javascript/api/excel/excel.identity#id)|Represents the user's unique ID.|
|[IdentityCollection](/javascript/api/excel/excel.identitycollection)|[add(assignee: Identity)](/javascript/api/excel/excel.identitycollection#add-assignee-)|Adds a user identity to the collection.|
||[clear()](/javascript/api/excel/excel.identitycollection#clear--)|Removes all user identities from the collection.|
||[getCount()](/javascript/api/excel/excel.identitycollection#getcount--)|Gets the number of items in the collection.|
||[getItemAt(index: number)](/javascript/api/excel/excel.identitycollection#getitemat-index-)|Gets a document user identity by using its index in the collection.|
||[remove(assignee: Identity)](/javascript/api/excel/excel.identitycollection#remove-assignee-)|Removes a user identity from the collection.|
|[InsertWorksheetOptions](/javascript/api/excel/excel.insertworksheetoptions)|[positionType](/javascript/api/excel/excel.insertworksheetoptions#positiontype)|The insert position, in the current workbook, of the new worksheets.|
||[relativeTo](/javascript/api/excel/excel.insertworksheetoptions#relativeto)|The worksheet in the current workbook that is referenced for the `WorksheetPositionType` parameter.|
||[sheetNamesToInsert](/javascript/api/excel/excel.insertworksheetoptions#sheetnamestoinsert)|The names of individual worksheets to insert.|
|[LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|[dataProvider](/javascript/api/excel/excel.linkeddatatype#dataprovider)|The name of the data provider for the linked data type.|
||[lastRefreshed](/javascript/api/excel/excel.linkeddatatype#lastrefreshed)|The local time-zone date and time since the workbook was opened when the linked data type was last refreshed.|
||[name](/javascript/api/excel/excel.linkeddatatype#name)|The name of the linked data type.|
||[periodicRefreshInterval](/javascript/api/excel/excel.linkeddatatype#periodicrefreshinterval)|The frequency, in seconds, at which the linked data type is refreshed if `refreshMode` is set to "Periodic".|
||[refreshMode](/javascript/api/excel/excel.linkeddatatype#refreshmode)|The mechanism by which the data for the linked data type is retrieved.|
||[serviceId](/javascript/api/excel/excel.linkeddatatype#serviceid)|The unique ID of the linked data type.|
||[supportedRefreshModes](/javascript/api/excel/excel.linkeddatatype#supportedrefreshmodes)|Returns an array with all the refresh modes supported by the linked data type.|
||[requestRefresh()](/javascript/api/excel/excel.linkeddatatype#requestrefresh--)|Makes a request to refresh the linked data type.|
||[requestSetRefreshMode(refreshMode: Excel.LinkedDataTypeRefreshMode)](/javascript/api/excel/excel.linkeddatatype#requestsetrefreshmode-refreshmode-)|Makes a request to change the refresh mode for this linked data type.|
|[LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs)|[serviceId](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#serviceid)|The unique ID of the new linked data type.|
||[source](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#source)|Gets the source of the event.|
||[type](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#type)|Gets the type of the event.|
|[LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection)|[getCount()](/javascript/api/excel/excel.linkeddatatypecollection#getcount--)|Gets the number of linked data types in the collection.|
||[getItem(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitem-key-)|Gets a linked data type by service ID.|
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
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getitemornullobject-key-)|Gets a sheet view using its name.|
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
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getFirstOrNullObject()](/javascript/api/excel/excel.pivottablescopedcollection#getfirstornullobject--)|Gets the first PivotTable in the collection.|
|[Range](/javascript/api/excel/excel.range)|[getDependents()](/javascript/api/excel/excel.range#getdependents--)|Returns a `WorkbookRangeAreas` object that represents the range containing all the dependents of a cell in the same worksheet or in multiple worksheets.|
||[getDirectDependents()](/javascript/api/excel/excel.range#getdirectdependents--)|Returns a `WorkbookRangeAreas` object that represents the range containing all the direct dependents of a cell in the same worksheet or in multiple worksheets.|
||[getExtendedRange(direction: Excel.KeyboardDirection, activeCell?: Range \| string)](/javascript/api/excel/excel.range#getextendedrange-direction--activecell-)|Returns a range object that includes the current range and up to the edge of the range, based on the provided direction.|
||[getMergedAreasOrNullObject()](/javascript/api/excel/excel.range#getmergedareasornullobject--)|Returns a RangeAreas object that represents the merged areas in this range.|
||[getPrecedents()](/javascript/api/excel/excel.range#getprecedents--)|Returns a `WorkbookRangeAreas` object that represents the range containing all the precedents of a cell in the same worksheet or in multiple worksheets.|
||[getRangeEdge(direction: Excel.KeyboardDirection, activeCell?: Range \| string)](/javascript/api/excel/excel.range#getrangeedge-direction--activecell-)|Returns a range object that is the edge cell of the data region that corresponds to the provided direction.|
|[Range](/javascript/api/excel/excel.range)||[RangeCustom](/javascript/api/excel/excel.rangecustom)||[RefreshModeChangedEventArgs](/javascript/api/excel/excel.refreshmodechangedeventargs)|[refreshMode](/javascript/api/excel/excel.refreshmodechangedeventargs#refreshmode)|The linked data type refresh mode.|
||[serviceId](/javascript/api/excel/excel.refreshmodechangedeventargs#serviceid)|The unique ID of the object whose refresh mode was changed.|
||[source](/javascript/api/excel/excel.refreshmodechangedeventargs#source)|Gets the source of the event.|
||[type](/javascript/api/excel/excel.refreshmodechangedeventargs#type)|Gets the type of the event.|
|[RefreshRequestCompletedEventArgs](/javascript/api/excel/excel.refreshrequestcompletedeventargs)|[refreshed](/javascript/api/excel/excel.refreshrequestcompletedeventargs#refreshed)|Indicates if the request to refresh was successful.|
||[serviceId](/javascript/api/excel/excel.refreshrequestcompletedeventargs#serviceid)|The unique ID of the object whose refresh request was completed.|
||[source](/javascript/api/excel/excel.refreshrequestcompletedeventargs#source)|Gets the source of the event.|
||[type](/javascript/api/excel/excel.refreshrequestcompletedeventargs#type)|Gets the type of the event.|
||[warnings](/javascript/api/excel/excel.refreshrequestcompletedeventargs#warnings)|An array that contains any warnings generated from the refresh request.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|Creates a scalable vector graphic (SVG) from an XML string and adds it to the worksheet.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.shapecollection#getitemornullobject-key-)|Gets a shape using its name or ID.|
|[Slicer](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|Represents the slicer name used in the formula.|
||[slicerStyle](/javascript/api/excel/excel.slicer#slicerstyle)|The style applied to the slicer.|
||[setStyle(style: string \| SlicerStyle \| BuiltInSlicerStyle)](/javascript/api/excel/excel.slicer#setstyle-style-)|Sets the style applied to the slicer.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getItemOrNullObject(name: string)](/javascript/api/excel/excel.stylecollection#getitemornullobject-name-)|Gets a style by name.|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|Changes the table to use the default table style.|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|Occurs when a filter is applied on a specific table.|
||[tableStyle](/javascript/api/excel/excel.table#tablestyle)|The style applied to the table.|
||[resize(newRange: Range \| string)](/javascript/api/excel/excel.table#resize-newrange-)|Resize the table to the new range.|
||[setStyle(style: string \| TableStyle \| BuiltInTableStyle)](/javascript/api/excel/excel.table#setstyle-style-)|Sets the style applied to the table.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|Occurs when a filter is applied on any table in a workbook, or a worksheet.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)||[TableCollectionCustom](/javascript/api/excel/excel.tablecollectioncustom)||[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|Gets the ID of the table in which the filter is applied.|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|Gets the ID of the worksheet which contains the table.|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.tablescopedcollection#getitemornullobject-key-)|Gets a table by name or ID.|
|[Workbook](/javascript/api/excel/excel.workbook)|[insertWorksheetsFromBase64(base64File: string, options?: Excel.InsertWorksheetOptions)](/javascript/api/excel/excel.workbook#insertworksheetsfrombase64-base64file--options-)|Inserts the specified worksheets from a source workbook into the current workbook.|
||[linkedDataTypes](/javascript/api/excel/excel.workbook#linkeddatatypes)|Returns a collection of linked data types that are part of the workbook.|
||[onActivated](/javascript/api/excel/excel.workbook#onactivated)|Occurs when the the workbook is activated.|
||[tasks](/javascript/api/excel/excel.workbook#tasks)|Returns a collection of tasks that are present in the workbook.|
||[showPivotFieldList](/javascript/api/excel/excel.workbook#showpivotfieldlist)|Specifies whether the PivotTable's field list pane is shown at the workbook level.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|True if the workbook uses the 1904 date system.|
|[WorkbookActivatedEventArgs](/javascript/api/excel/excel.workbookactivatedeventargs)|[type](/javascript/api/excel/excel.workbookactivatedeventargs#type)|Gets the type of the event.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[namedSheetViews](/javascript/api/excel/excel.worksheet#namedsheetviews)|Returns a collection of sheet views that are present in the worksheet.|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|Occurs when a filter is applied on a specific worksheet.|
||[onFormulaChanged](/javascript/api/excel/excel.worksheet#onformulachanged)|Occurs when one or more formulas are changed in this worksheet.|
||[tasks](/javascript/api/excel/excel.worksheet#tasks)|Returns a collection of tasks that are present in the worksheet.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|Inserts the specified worksheets of a workbook into the current workbook.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|Occurs when any worksheet's filter is applied in the workbook.|
||[onFormulaChanged](/javascript/api/excel/excel.worksheetcollection#onformulachanged)|Occurs when one or more formulas are changed in any worksheet of this collection.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|Gets the ID of the worksheet in which the filter is applied.|
|[WorksheetFormulaChangedEventArgs](/javascript/api/excel/excel.worksheetformulachangedeventargs)|[formulaDetails](/javascript/api/excel/excel.worksheetformulachangedeventargs#formuladetails)|Gets an array of `FormulaChangedEventDetail` objects, which contain the details about the all of the changed formulas.|
||[source](/javascript/api/excel/excel.worksheetformulachangedeventargs#source)|The source of the event.|
||[type](/javascript/api/excel/excel.worksheetformulachangedeventargs#type)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetformulachangedeventargs#worksheetid)|Gets the ID of the worksheet in which the formula changed.|

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel?view=excel-js-preview&preserve-view=true)
- [Excel JavaScript API requirement sets](excel-api-requirement-sets.md)
