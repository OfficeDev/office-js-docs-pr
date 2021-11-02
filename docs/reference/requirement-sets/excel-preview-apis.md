---
title: Excel JavaScript preview APIs
description: 'Details about upcoming Excel JavaScript APIs.'
ms.date: 11/02/2021
ms.prod: excel
ms.localizationpriority: medium
---

# Excel JavaScript preview APIs

New Excel JavaScript APIs are first introduced in "preview" and later become part of a specific, numbered requirement set after sufficient testing occurs and user feedback is acquired.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

The following table provides a concise summary of the APIs, while the subsequent [API list](#api-list) table gives a detailed list.

| Feature area | Description | Relevant objects |
|:--- |:--- |:--- |
| [Data types](../../excel/excel-data-types-overview.md) | An extension of existing Excel data types, including support for formatted numbers and web images. | [ArrayCellValue](/javascript/api/excel/excel.arraycellvalue), [BooleanCellValue](/javascript/api/excel/excel.booleancellvalue), [CellValueAttributionAttributes](/javascript/api/excel/excel.cellvalueattributionattributes), [CellValueProviderAttributes](/javascript/api/excel/excel.cellvalueproviderattributes), [DoubleCellValue](/javascript/api/excel/excel.doublecellvalue), [EmptyCellValue](/javascript/api/excel/excel.emptycellvalue), [EntityCellValue](/javascript/api/excel/excel.entitycellvalue), [FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue), [StringCellValue](/javascript/api/excel/excel.stringcellvalue), [ValueTypeNotAvailableCellValue](/javascript/api/excel/excel.valuetypenotavailablecellvalue), [WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue) |
| [Data types errors](../../excel/excel-data-types-concepts.md#improved-error-support) | Error objects that support expanded data types. | [BlockedErrorCellValue](/javascript/api/excel/excel.blockederrorcellvalue), [BusyErrorCellValue](/javascript/api/excel/excel.busyerrorcellvalue), [CalcErrorCellValue](/javascript/api/excel/excel.calcerrorcellvalue), [ConnectErrorCellValue](/javascript/api/excel/excel.connecterrorcellvalue), [Div0ErrorCellValue](/javascript/api/excel/excel.div0errorcellvalue), [FieldErrorCellValue](/javascript/api/excel/excel.fielderrorcellvalue), [GettingDataErrorCellValue](/javascript/api/excel/excel.gettingdataerrorcellvalue), [NotAvailableErrorCellValue](/javascript/api/excel/excel.notavailableerrorcellvalue), [NameErrorCellValue](/javascript/api/excel/excel.nameerrorcellvalue), [NullErrorCellValue](/javascript/api/excel/excel.nullerrorcellvalue), [NumErrorCellValue](/javascript/api/excel/excel.numerrorcellvalue), [RefErrorCellValue](/javascript/api/excel/excel.referrorcellvalue), [SpillErrorCellValue](/javascript/api/excel/excel.spillerrorcellvalue), [ValueErrorCellValue](/javascript/api/excel/excel.valueerrorcellvalue)|
| Document tasks | Turn comments into tasks assigned to users. | [DocumentTask](/javascript/api/excel/excel.documenttask) |
| Identities | Manage user identities, including display name and email address. | [Identity](/javascript/api/excel/excel.identity), [IdentityCollection](/javascript/api/excel/excel.identitycollection), [IdentityEntity](/javascript/api/excel/excel.identityentity) |
| Linked data types | Adds support for data types connected to Excel from external sources. | [LinkedDataType](/javascript/api/excel/excel.linkeddatatype), [LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs), [LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection) |
| Table styles | Provides control for font, border, fill color, and other aspects of table styles. | [Table](/javascript/api/excel/excel.table), [PivotTable](/javascript/api/excel/excel.pivottable), [Slicer](/javascript/api/excel/excel.slicer) |
| Worksheet protection | Prevent unauthorized users from making changes to specified ranges within a worksheet. | [WorksheetProtection](/javascript/api/excel/excel.worksheetprotection), [AllowEditRange](/javascript/api/excel/excel.alloweditrange), [AllowEditRangeCollection](/javascript/api/excel/excel.alloweditrangecollection), [AllowEditRangeOptions](/javascript/api/excel/excel.alloweditrangeoptions) |

## API list

The following table lists the Excel JavaScript APIs currently in preview. For a complete list of all Excel JavaScript APIs (including preview APIs and previously released APIs), see [all Excel JavaScript APIs](/javascript/api/excel?view=excel-js-preview&preserve-view=true).

| Class | Fields | Description |
|:---|:---|:---|
|[AllowEditRange](/javascript/api/excel/excel.alloweditrange)|[address](/javascript/api/excel/excel.alloweditrange#address)|Specifies the range associated with the object.|
||[delete()](/javascript/api/excel/excel.alloweditrange#delete__)|Deletes this object from the `AllowEditRangeCollection`.|
||[isPasswordProtected](/javascript/api/excel/excel.alloweditrange#isPasswordProtected)|Specifies if the `AllowEditRange` is password protected.|
||[pauseProtection(password?: string)](/javascript/api/excel/excel.alloweditrange#pauseProtection_password_)|Pauses worksheet protection for the given `AllowEditRange` object for the user in a given session.|
||[setPassword(password?: string)](/javascript/api/excel/excel.alloweditrange#setPassword_password_)|Changes the password associated with the `AllowEditRange`.|
||[title](/javascript/api/excel/excel.alloweditrange#title)|Specifies the title of the object.|
|[AllowEditRangeCollection](/javascript/api/excel/excel.alloweditrangecollection)|[add(title: string, rangeAddress: string, options?: Excel.AllowEditRangeOptions)](/javascript/api/excel/excel.alloweditrangecollection#add_title__rangeAddress__options_)|Adds an `AllowEditRange` object to the collection.|
||[getCount()](/javascript/api/excel/excel.alloweditrangecollection#getCount__)|Returns the number of `AllowEditRange` objects in the collection.|
||[getItem(key: string)](/javascript/api/excel/excel.alloweditrangecollection#getItem_key_)|Gets the `AllowEditRange` object by its title.|
||[getItemAt(index: number)](/javascript/api/excel/excel.alloweditrangecollection#getItemAt_index_)|Returns an `AllowEditRange` object by its index in the collection.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.alloweditrangecollection#getItemOrNullObject_key_)|Gets the `AllowEditRange` object by its title.|
||[items](/javascript/api/excel/excel.alloweditrangecollection#items)|Gets the loaded child items in this collection.|
||[pauseProtection(password: string)](/javascript/api/excel/excel.alloweditrangecollection#pauseProtection_password_)|Pauses worksheet protection for all `AllowEditRange` objects in the collection that have the given password for the user in a given session.|
|[AllowEditRangeOptions](/javascript/api/excel/excel.alloweditrangeoptions)|[password](/javascript/api/excel/excel.alloweditrangeoptions#password)|The password associated with the `AllowEditRange`.|
|[ArrayCellValue](/javascript/api/excel/excel.arraycellvalue)|[basicType](/javascript/api/excel/excel.arraycellvalue#basicType)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.arraycellvalue#basicValue)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[elements](/javascript/api/excel/excel.arraycellvalue#elements)|Represents the elements of the array.|
||[type](/javascript/api/excel/excel.arraycellvalue#type)|Represents the type of this cell value.|
|[BlockedErrorCellValue](/javascript/api/excel/excel.blockederrorcellvalue)|[basicType](/javascript/api/excel/excel.blockederrorcellvalue#basicType)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.blockederrorcellvalue#basicValue)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[errorSubType](/javascript/api/excel/excel.blockederrorcellvalue#errorSubType)|Represents the type of `BlockedErrorCellValue`.|
||[errorType](/javascript/api/excel/excel.blockederrorcellvalue#errorType)|Represents the type of `ErrorCellValue`.|
||[type](/javascript/api/excel/excel.blockederrorcellvalue#type)|Represents the type of this cell value.|
|[BooleanCellValue](/javascript/api/excel/excel.booleancellvalue)|[basicType](/javascript/api/excel/excel.booleancellvalue#basicType)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.booleancellvalue#basicValue)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[type](/javascript/api/excel/excel.booleancellvalue#type)|Represents the type of this cell value.|
|[BusyErrorCellValue](/javascript/api/excel/excel.busyerrorcellvalue)|[basicType](/javascript/api/excel/excel.busyerrorcellvalue#basicType)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.busyerrorcellvalue#basicValue)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[errorSubType](/javascript/api/excel/excel.busyerrorcellvalue#errorSubType)|Represents the type of `BusyErrorCellValue`.|
||[errorType](/javascript/api/excel/excel.busyerrorcellvalue#errorType)|Represents the type of `ErrorCellValue`.|
||[type](/javascript/api/excel/excel.busyerrorcellvalue#type)|Represents the type of this cell value.|
|[CalcErrorCellValue](/javascript/api/excel/excel.calcerrorcellvalue)|[basicType](/javascript/api/excel/excel.calcerrorcellvalue#basicType)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.calcerrorcellvalue#basicValue)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[errorSubType](/javascript/api/excel/excel.calcerrorcellvalue#errorSubType)|Represents the type of `CalcErrorCellValue`.|
||[errorType](/javascript/api/excel/excel.calcerrorcellvalue#errorType)|Represents the type of `ErrorCellValue`.|
||[type](/javascript/api/excel/excel.calcerrorcellvalue#type)|Represents the type of this cell value.|
|[CellValueAttributionAttributes](/javascript/api/excel/excel.cellvalueattributionattributes)|[licenseAddress](/javascript/api/excel/excel.cellvalueattributionattributes#licenseAddress)|Represents a URL to a license or source that describes how this property can be used.|
||[licenseText](/javascript/api/excel/excel.cellvalueattributionattributes#licenseText)|Represents a name for the license that governs this property.|
||[sourceAddress](/javascript/api/excel/excel.cellvalueattributionattributes#sourceAddress)|Represents a URL to the source of the `CellValue`.|
||[sourceText](/javascript/api/excel/excel.cellvalueattributionattributes#sourceText)|Represents a name for the source of the `CellValue`.|
|[CellValuePropertyMetadata](/javascript/api/excel/excel.cellvaluepropertymetadata)|[attribution](/javascript/api/excel/excel.cellvaluepropertymetadata#attribution)|Represents attribution information to describe the source and license requirements for using this property.|
||[excludeFrom](/javascript/api/excel/excel.cellvaluepropertymetadata#excludeFrom)|Represents which features this property is excluded from.|
||[sublabel](/javascript/api/excel/excel.cellvaluepropertymetadata#sublabel)|Represents the sublabel for this property shown in card view.|
|[CellValuePropertyMetadataExclusions](/javascript/api/excel/excel.cellvaluepropertymetadataexclusions)|[autoComplete](/javascript/api/excel/excel.cellvaluepropertymetadataexclusions#autoComplete)|True represents that the property is excluded from the properties shown by auto complete.|
||[calcCompare](/javascript/api/excel/excel.cellvaluepropertymetadataexclusions#calcCompare)|True represents that the property is excluded from the properties used to compare cell values during recalc.|
||[cardView](/javascript/api/excel/excel.cellvaluepropertymetadataexclusions#cardView)|True represents that the property is excluded from the properties shown by card view.|
||[dotNotation](/javascript/api/excel/excel.cellvaluepropertymetadataexclusions#dotNotation)|True represents that the property is excluded from the properties which can be accessed via the FIELDVALUE function.|
|[CellValueProviderAttributes](/javascript/api/excel/excel.cellvalueproviderattributes)|[description](/javascript/api/excel/excel.cellvalueproviderattributes#description)|Represents the provider description property that is used in card view if no logo is specified.|
||[logoSourceAddress](/javascript/api/excel/excel.cellvalueproviderattributes#logoSourceAddress)|Represents a URL used to download an image that will be used as a logo in card view.|
||[logoTargetAddress](/javascript/api/excel/excel.cellvalueproviderattributes#logoTargetAddress)|Represents a URL that is the navigation target if the user clicks on the logo element in card view.|
|[Comment](/javascript/api/excel/excel.comment)|[assignTask(assignee: Identity)](/javascript/api/excel/excel.comment#assignTask_assignee_)|Assigns the task attached to the comment to the given user as an assignee.|
||[getTask()](/javascript/api/excel/excel.comment#getTask__)|Gets the task associated with this comment.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.comment#getTaskOrNullObject__)|Gets the task associated with this comment.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[assignTask(assignee: Identity)](/javascript/api/excel/excel.commentreply#assignTask_assignee_)|Assigns the task attached to the comment to the given user as the sole assignee.|
||[getTask()](/javascript/api/excel/excel.commentreply#getTask__)|Gets the task associated with this comment reply's thread.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.commentreply#getTaskOrNullObject__)|Gets the task associated with this comment reply's thread.|
|[ConnectErrorCellValue](/javascript/api/excel/excel.connecterrorcellvalue)|[basicType](/javascript/api/excel/excel.connecterrorcellvalue#basicType)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.connecterrorcellvalue#basicValue)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[errorSubType](/javascript/api/excel/excel.connecterrorcellvalue#errorSubType)|Represents the type of `ConnectErrorCellValue`.|
||[errorType](/javascript/api/excel/excel.connecterrorcellvalue#errorType)|Represents the type of `ErrorCellValue`.|
||[type](/javascript/api/excel/excel.connecterrorcellvalue#type)|Represents the type of this cell value.|
|[Div0ErrorCellValue](/javascript/api/excel/excel.div0errorcellvalue)|[basicType](/javascript/api/excel/excel.div0errorcellvalue#basicType)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.div0errorcellvalue#basicValue)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[errorType](/javascript/api/excel/excel.div0errorcellvalue#errorType)|Represents the type of `ErrorCellValue`.|
||[type](/javascript/api/excel/excel.div0errorcellvalue#type)|Represents the type of this cell value.|
|[DocumentTask](/javascript/api/excel/excel.documenttask)|[assignees](/javascript/api/excel/excel.documenttask#assignees)|Returns a collection of assignees of the task.|
||[changes](/javascript/api/excel/excel.documenttask#changes)|Gets the change records of the task.|
||[comment](/javascript/api/excel/excel.documenttask#comment)|Gets the comment associated with the task.|
||[completedBy](/javascript/api/excel/excel.documenttask#completedBy)|Gets the most recent user to have completed the task.|
||[completedDateTime](/javascript/api/excel/excel.documenttask#completedDateTime)|Gets the date and time that the task was completed.|
||[createdBy](/javascript/api/excel/excel.documenttask#createdBy)|Gets the user who created the task.|
||[createdDateTime](/javascript/api/excel/excel.documenttask#createdDateTime)|Gets the date and time that the task was created.|
||[id](/javascript/api/excel/excel.documenttask#id)|Gets the ID of the task.|
||[percentComplete](/javascript/api/excel/excel.documenttask#percentComplete)|Specifies the completion percentage of the task.|
||[priority](/javascript/api/excel/excel.documenttask#priority)|Specifies the priority of the task.|
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
|[DoubleCellValue](/javascript/api/excel/excel.doublecellvalue)|[basicType](/javascript/api/excel/excel.doublecellvalue#basicType)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.doublecellvalue#basicValue)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[type](/javascript/api/excel/excel.doublecellvalue#type)|Represents the type of this cell value.|
|[EmptyCellValue](/javascript/api/excel/excel.emptycellvalue)|[basicType](/javascript/api/excel/excel.emptycellvalue#basicType)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.emptycellvalue#basicValue)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[type](/javascript/api/excel/excel.emptycellvalue#type)|Represents the type of this cell value.|
|[EntityCellValue](/javascript/api/excel/excel.entitycellvalue)|[basicType](/javascript/api/excel/excel.entitycellvalue#basicType)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.entitycellvalue#basicValue)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[properties: {            [key: string]: CellValue & {                propertyMetadata](/javascript/api/excel/excel.entitycellvalue#properties)|Represents the properties of this entity and their metadata.|
||[propertyMetadata](/javascript/api/excel/excel.entitycellvalue#propertyMetadata)||
||[text](/javascript/api/excel/excel.entitycellvalue#text)|Represents the text shown when a cell with this value is rendered.|
||[type](/javascript/api/excel/excel.entitycellvalue#type)|Represents the type of this cell value.|
|[FieldErrorCellValue](/javascript/api/excel/excel.fielderrorcellvalue)|[basicType](/javascript/api/excel/excel.fielderrorcellvalue#basicType)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.fielderrorcellvalue#basicValue)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[errorSubType](/javascript/api/excel/excel.fielderrorcellvalue#errorSubType)|Represents the type of `FieldErrorCellValue`.|
||[errorType](/javascript/api/excel/excel.fielderrorcellvalue#errorType)|Represents the type of `ErrorCellValue`.|
||[type](/javascript/api/excel/excel.fielderrorcellvalue#type)|Represents the type of this cell value.|
|[FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue)|[basicType](/javascript/api/excel/excel.formattednumbercellvalue#basicType)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.formattednumbercellvalue#basicValue)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[numberFormat](/javascript/api/excel/excel.formattednumbercellvalue#numberFormat)|Returns the number format string that is used to display this value.|
||[type](/javascript/api/excel/excel.formattednumbercellvalue#type)|Represents the type of this cell value.|
|[GettingDataErrorCellValue](/javascript/api/excel/excel.gettingdataerrorcellvalue)|[basicType](/javascript/api/excel/excel.gettingdataerrorcellvalue#basicType)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.gettingdataerrorcellvalue#basicValue)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[errorType](/javascript/api/excel/excel.gettingdataerrorcellvalue#errorType)|Represents the type of `ErrorCellValue`.|
||[type](/javascript/api/excel/excel.gettingdataerrorcellvalue#type)|Represents the type of this cell value.|
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
||[requestRefresh()](/javascript/api/excel/excel.linkeddatatype#requestRefresh__)|Makes a request to refresh the linked data type.|
||[requestSetRefreshMode(refreshMode: Excel.LinkedDataTypeRefreshMode)](/javascript/api/excel/excel.linkeddatatype#requestSetRefreshMode_refreshMode_)|Makes a request to change the refresh mode for this linked data type.|
||[serviceId](/javascript/api/excel/excel.linkeddatatype#serviceId)|The unique ID of the linked data type.|
||[supportedRefreshModes](/javascript/api/excel/excel.linkeddatatype#supportedRefreshModes)|Returns an array with all the refresh modes supported by the linked data type.|
|[LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs)|[serviceId](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#serviceId)|The unique ID of the new linked data type.|
||[source](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#source)|Gets the source of the event.|
||[type](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#type)|Gets the type of the event.|
|[LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection)|[getCount()](/javascript/api/excel/excel.linkeddatatypecollection#getCount__)|Gets the number of linked data types in the collection.|
||[getItem(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getItem_key_)|Gets a linked data type by service ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.linkeddatatypecollection#getItemAt_index_)|Gets a linked data type by its index in the collection.|
||[getItemOrNullObject(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getItemOrNullObject_key_)|Gets a linked data type by ID.|
||[items](/javascript/api/excel/excel.linkeddatatypecollection#items)|Gets the loaded child items in this collection.|
||[requestRefreshAll()](/javascript/api/excel/excel.linkeddatatypecollection#requestRefreshAll__)|Makes a request to refresh all the linked data types in the collection.|
|[NameErrorCellValue](/javascript/api/excel/excel.nameerrorcellvalue)|[basicType](/javascript/api/excel/excel.nameerrorcellvalue#basicType)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.nameerrorcellvalue#basicValue)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[errorType](/javascript/api/excel/excel.nameerrorcellvalue#errorType)|Represents the type of `ErrorCellValue`.|
||[type](/javascript/api/excel/excel.nameerrorcellvalue#type)|Represents the type of this cell value.|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getItemOrNullObject_key_)|Gets a sheet view using its name.|
|[NotAvailableErrorCellValue](/javascript/api/excel/excel.notavailableerrorcellvalue)|[basicType](/javascript/api/excel/excel.notavailableerrorcellvalue#basicType)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.notavailableerrorcellvalue#basicValue)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[errorType](/javascript/api/excel/excel.notavailableerrorcellvalue#errorType)|Represents the type of `ErrorCellValue`.|
||[type](/javascript/api/excel/excel.notavailableerrorcellvalue#type)|Represents the type of this cell value.|
|[NullErrorCellValue](/javascript/api/excel/excel.nullerrorcellvalue)|[basicType](/javascript/api/excel/excel.nullerrorcellvalue#basicType)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.nullerrorcellvalue#basicValue)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[errorType](/javascript/api/excel/excel.nullerrorcellvalue#errorType)|Represents the type of `ErrorCellValue`.|
||[type](/javascript/api/excel/excel.nullerrorcellvalue#type)|Represents the type of this cell value.|
|[NumErrorCellValue](/javascript/api/excel/excel.numerrorcellvalue)|[basicType](/javascript/api/excel/excel.numerrorcellvalue#basicType)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.numerrorcellvalue#basicValue)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[errorType](/javascript/api/excel/excel.numerrorcellvalue#errorType)|Represents the type of `ErrorCellValue`.|
||[type](/javascript/api/excel/excel.numerrorcellvalue#type)|Represents the type of this cell value.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getCell_dataHierarchy__rowItems__columnItems_)|Gets a unique cell in the PivotTable based on a data hierarchy and the row and column items of their respective hierarchies.|
||[pivotStyle](/javascript/api/excel/excel.pivotlayout#pivotStyle)|The style applied to the PivotTable.|
||[setStyle(style: string \| PivotTableStyle \| BuiltInPivotTableStyle)](/javascript/api/excel/excel.pivotlayout#setStyle_style_)|Sets the style applied to the PivotTable.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[getDataSourceString()](/javascript/api/excel/excel.pivottable#getDataSourceString__)|Returns the string representation of the data source for the PivotTable.|
||[getDataSourceType()](/javascript/api/excel/excel.pivottable#getDataSourceType__)|Gets the type of the data source for the PivotTable.|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getFirstOrNullObject()](/javascript/api/excel/excel.pivottablescopedcollection#getFirstOrNullObject__)|Gets the first PivotTable in the collection.|
|[Range](/javascript/api/excel/excel.range)|[getDependents()](/javascript/api/excel/excel.range#getDependents__)|Returns a `WorkbookRangeAreas` object that represents the range containing all the dependents of a cell in the same worksheet or in multiple worksheets.|
|[RefErrorCellValue](/javascript/api/excel/excel.referrorcellvalue)|[basicType](/javascript/api/excel/excel.referrorcellvalue#basicType)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.referrorcellvalue#basicValue)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[errorSubType](/javascript/api/excel/excel.referrorcellvalue#errorSubType)|Represents the type of `RefErrorCellValue`.|
||[errorType](/javascript/api/excel/excel.referrorcellvalue#errorType)|Represents the type of `ErrorCellValue`.|
||[type](/javascript/api/excel/excel.referrorcellvalue#type)|Represents the type of this cell value.|
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
|[Slicer](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameInFormula)|Represents the slicer name used in the formula.|
||[setStyle(style: string \| SlicerStyle \| BuiltInSlicerStyle)](/javascript/api/excel/excel.slicer#setStyle_style_)|Sets the style applied to the slicer.|
||[slicerStyle](/javascript/api/excel/excel.slicer#slicerStyle)|The style applied to the slicer.|
|[SpillErrorCellValue](/javascript/api/excel/excel.spillerrorcellvalue)|[basicType](/javascript/api/excel/excel.spillerrorcellvalue#basicType)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.spillerrorcellvalue#basicValue)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[errorSubType](/javascript/api/excel/excel.spillerrorcellvalue#errorSubType)|Represents the type of `SpillErrorCellValue`.|
||[errorType](/javascript/api/excel/excel.spillerrorcellvalue#errorType)|Represents the type of `ErrorCellValue`.|
||[spilledColumns](/javascript/api/excel/excel.spillerrorcellvalue#spilledColumns)|Represents the number of columns that would spill if there were no #SPILL! error.|
||[spilledRows](/javascript/api/excel/excel.spillerrorcellvalue#spilledRows)|Represents the number of rows that would spill if there were no #SPILL! error.|
||[type](/javascript/api/excel/excel.spillerrorcellvalue#type)|Represents the type of this cell value.|
|[StringCellValue](/javascript/api/excel/excel.stringcellvalue)|[basicType](/javascript/api/excel/excel.stringcellvalue#basicType)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.stringcellvalue#basicValue)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[type](/javascript/api/excel/excel.stringcellvalue#type)|Represents the type of this cell value.|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearStyle__)|Changes the table to use the default table style.|
||[onFiltered](/javascript/api/excel/excel.table#onFiltered)|Occurs when a filter is applied on a specific table.|
||[setStyle(style: string \| TableStyle \| BuiltInTableStyle)](/javascript/api/excel/excel.table#setStyle_style_)|Sets the style applied to the table.|
||[tableStyle](/javascript/api/excel/excel.table#tableStyle)|The style applied to the table.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onFiltered)|Occurs when a filter is applied on any table in a workbook, or a worksheet.|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableId)|Gets the ID of the table in which the filter is applied.|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetId)|Gets the ID of the worksheet which contains the table.|
|[ValueErrorCellValue](/javascript/api/excel/excel.valueerrorcellvalue)|[basicType](/javascript/api/excel/excel.valueerrorcellvalue#basicType)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.valueerrorcellvalue#basicValue)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[errorSubType](/javascript/api/excel/excel.valueerrorcellvalue#errorSubType)|Represents the type of `ValueErrorCellValue`.|
||[errorType](/javascript/api/excel/excel.valueerrorcellvalue#errorType)|Represents the type of `ErrorCellValue`.|
||[type](/javascript/api/excel/excel.valueerrorcellvalue#type)|Represents the type of this cell value.|
|[ValueTypeNotAvailableCellValue](/javascript/api/excel/excel.valuetypenotavailablecellvalue)|[basicType](/javascript/api/excel/excel.valuetypenotavailablecellvalue#basicType)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.valuetypenotavailablecellvalue#basicValue)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[type](/javascript/api/excel/excel.valuetypenotavailablecellvalue#type)|Represents the type of this cell value.|
|[WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue)|[address](/javascript/api/excel/excel.webimagecellvalue#address)|Represents the URL from which the image will be downloaded.|
||[altText](/javascript/api/excel/excel.webimagecellvalue#altText)|Represents the alternate text that can be used in accessibility scenarios to describe what the image represents.|
||[attribution](/javascript/api/excel/excel.webimagecellvalue#attribution)|Represents attribution information to describe the source and license requirements for using this image.|
||[basicType](/javascript/api/excel/excel.webimagecellvalue#basicType)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.webimagecellvalue#basicValue)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[provider](/javascript/api/excel/excel.webimagecellvalue#provider)|Represents information that describes the entity or individual who provided the image.|
||[relatedImagesAddress](/javascript/api/excel/excel.webimagecellvalue#relatedImagesAddress)|Represents the URL of a webpage with images that are considered related to this `WebImageCellValue`.|
||[type](/javascript/api/excel/excel.webimagecellvalue#type)|Represents the type of this cell value.|
|[Workbook](/javascript/api/excel/excel.workbook)|[linkedDataTypes](/javascript/api/excel/excel.workbook#linkedDataTypes)|Returns a collection of linked data types that are part of the workbook.|
||[showPivotFieldList](/javascript/api/excel/excel.workbook#showPivotFieldList)|Specifies whether the PivotTable's field list pane is shown at the workbook level.|
||[tasks](/javascript/api/excel/excel.workbook#tasks)|Returns a collection of tasks that are present in the workbook.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904DateSystem)|True if the workbook uses the 1904 date system.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onFiltered](/javascript/api/excel/excel.worksheet#onFiltered)|Occurs when a filter is applied on a specific worksheet.|
||[tasks](/javascript/api/excel/excel.worksheet#tasks)|Returns a collection of tasks that are present in the worksheet.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addFromBase64_base64File__sheetNamesToInsert__positionType__relativeTo_)|Inserts the specified worksheets of a workbook into the current workbook.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onFiltered)|Occurs when any worksheet's filter is applied in the workbook.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetId)|Gets the ID of the worksheet in which the filter is applied.|
|[WorksheetProtection](/javascript/api/excel/excel.worksheetprotection)|[allowEditRanges](/javascript/api/excel/excel.worksheetprotection#allowEditRanges)|Specifies the `AllowEditRangeCollection` found in this worksheet.|
||[canPauseProtection](/javascript/api/excel/excel.worksheetprotection#canPauseProtection)|Specifies if protection can be paused for this worksheet.|
||[checkPassword(password?: string)](/javascript/api/excel/excel.worksheetprotection#checkPassword_password_)|Specifies if the password can be used to unlock worksheet protection.|
||[isPasswordProtected](/javascript/api/excel/excel.worksheetprotection#isPasswordProtected)|Specifies if the sheet is password protected.|
||[isPaused](/javascript/api/excel/excel.worksheetprotection#isPaused)|Specifies if worksheet protection is paused.|
||[pauseProtection(password?: string)](/javascript/api/excel/excel.worksheetprotection#pauseProtection_password_)|Pauses worksheet protection for the given worksheet object for the user in a given session.|
||[resumeProtection()](/javascript/api/excel/excel.worksheetprotection#resumeProtection__)|Resumes worksheet protection for the given worksheet object for the user in a given session.|
||[setPassword(password?: string)](/javascript/api/excel/excel.worksheetprotection#setPassword_password_)|Changes the password associated with the `WorksheetProtection` object.|
||[updateOptions(options: Excel.WorksheetProtectionOptions)](/javascript/api/excel/excel.worksheetprotection#updateOptions_options_)|Change the worksheet protection options associated to the `WorksheetProtection` object.|
|[WorksheetProtectionChangedEventArgs](/javascript/api/excel/excel.worksheetprotectionchangedeventargs)|[allowEditRangesChanged](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#allowEditRangesChanged)|Specifies if any of the `AllowEditRange` objects have changed.|
||[protectionOptionsChanged](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#protectionOptionsChanged)|Specifies if the `WorksheetProtectionOptions` have changed.|
||[sheetPasswordChanged](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#sheetPasswordChanged)|Specifies if the worksheet password has changed.|

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel?view=excel-js-preview&preserve-view=true)
- [Excel JavaScript API requirement sets](excel-api-requirement-sets.md)
