---
title: Excel JavaScript preview APIs
description: 'Details about upcoming Excel JavaScript APIs.'
ms.date: 12/08/2021
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
|[AllowEditRange](/javascript/api/excel/excel.alloweditrange)|[address](/javascript/api/excel/excel.alloweditrange#excel-excel-alloweditrange-address-member)|Specifies the range associated with the object.|
||[delete()](/javascript/api/excel/excel.alloweditrange#excel-excel-alloweditrange-delete-member(1))|Deletes this object from the `AllowEditRangeCollection`.|
||[isPasswordProtected](/javascript/api/excel/excel.alloweditrange#excel-excel-alloweditrange-ispasswordprotected-member)|Specifies if the `AllowEditRange` is password protected.|
||[pauseProtection(password?: string)](/javascript/api/excel/excel.alloweditrange#excel-excel-alloweditrange-pauseprotection-member(1))|Pauses worksheet protection for the given `AllowEditRange` object for the user in a given session.|
||[setPassword(password?: string)](/javascript/api/excel/excel.alloweditrange#excel-excel-alloweditrange-setpassword-member(1))|Changes the password associated with the `AllowEditRange`.|
||[title](/javascript/api/excel/excel.alloweditrange#excel-excel-alloweditrange-title-member)|Specifies the title of the object.|
|[AllowEditRangeCollection](/javascript/api/excel/excel.alloweditrangecollection)|[add(title: string, rangeAddress: string, options?: Excel.AllowEditRangeOptions)](/javascript/api/excel/excel.alloweditrangecollection#excel-excel-alloweditrangecollection-add-member(1))|Adds an `AllowEditRange` object to the collection.|
||[getCount()](/javascript/api/excel/excel.alloweditrangecollection#excel-excel-alloweditrangecollection-getcount-member(1))|Returns the number of `AllowEditRange` objects in the collection.|
||[getItem(key: string)](/javascript/api/excel/excel.alloweditrangecollection#excel-excel-alloweditrangecollection-getitem-member(1))|Gets the `AllowEditRange` object by its title.|
||[getItemAt(index: number)](/javascript/api/excel/excel.alloweditrangecollection#excel-excel-alloweditrangecollection-getitemat-member(1))|Returns an `AllowEditRange` object by its index in the collection.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.alloweditrangecollection#excel-excel-alloweditrangecollection-getitemornullobject-member(1))|Gets the `AllowEditRange` object by its title.|
||[items](/javascript/api/excel/excel.alloweditrangecollection#excel-excel-alloweditrangecollection-items-member)|Gets the loaded child items in this collection.|
||[pauseProtection(password: string)](/javascript/api/excel/excel.alloweditrangecollection#excel-excel-alloweditrangecollection-pauseprotection-member(1))|Pauses worksheet protection for all `AllowEditRange` objects in the collection that have the given password for the user in a given session.|
|[AllowEditRangeOptions](/javascript/api/excel/excel.alloweditrangeoptions)|[password](/javascript/api/excel/excel.alloweditrangeoptions#excel-excel-alloweditrangeoptions-password-member)|The password associated with the `AllowEditRange`.|
|[ArrayCellValue](/javascript/api/excel/excel.arraycellvalue)|[basicType](/javascript/api/excel/excel.arraycellvalue#excel-excel-arraycellvalue-basictype-member)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.arraycellvalue#excel-excel-arraycellvalue-basicvalue-member)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[elements](/javascript/api/excel/excel.arraycellvalue#excel-excel-arraycellvalue-elements-member)|Represents the elements of the array.|
||[type](/javascript/api/excel/excel.arraycellvalue#excel-excel-arraycellvalue-type-member)|Represents the type of this cell value.|
|[BlockedErrorCellValue](/javascript/api/excel/excel.blockederrorcellvalue)|[basicType](/javascript/api/excel/excel.blockederrorcellvalue#excel-excel-blockederrorcellvalue-basictype-member)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.blockederrorcellvalue#excel-excel-blockederrorcellvalue-basicvalue-member)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[errorSubType](/javascript/api/excel/excel.blockederrorcellvalue#excel-excel-blockederrorcellvalue-errorsubtype-member)|Represents the type of `BlockedErrorCellValue`.|
||[errorType](/javascript/api/excel/excel.blockederrorcellvalue#excel-excel-blockederrorcellvalue-errortype-member)|Represents the type of `ErrorCellValue`.|
||[type](/javascript/api/excel/excel.blockederrorcellvalue#excel-excel-blockederrorcellvalue-type-member)|Represents the type of this cell value.|
|[BooleanCellValue](/javascript/api/excel/excel.booleancellvalue)|[basicType](/javascript/api/excel/excel.booleancellvalue#excel-excel-booleancellvalue-basictype-member)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.booleancellvalue#excel-excel-booleancellvalue-basicvalue-member)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[type](/javascript/api/excel/excel.booleancellvalue#excel-excel-booleancellvalue-type-member)|Represents the type of this cell value.|
|[BusyErrorCellValue](/javascript/api/excel/excel.busyerrorcellvalue)|[basicType](/javascript/api/excel/excel.busyerrorcellvalue#excel-excel-busyerrorcellvalue-basictype-member)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.busyerrorcellvalue#excel-excel-busyerrorcellvalue-basicvalue-member)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[errorSubType](/javascript/api/excel/excel.busyerrorcellvalue#excel-excel-busyerrorcellvalue-errorsubtype-member)|Represents the type of `BusyErrorCellValue`.|
||[errorType](/javascript/api/excel/excel.busyerrorcellvalue#excel-excel-busyerrorcellvalue-errortype-member)|Represents the type of `ErrorCellValue`.|
||[type](/javascript/api/excel/excel.busyerrorcellvalue#excel-excel-busyerrorcellvalue-type-member)|Represents the type of this cell value.|
|[CalcErrorCellValue](/javascript/api/excel/excel.calcerrorcellvalue)|[basicType](/javascript/api/excel/excel.calcerrorcellvalue#excel-excel-calcerrorcellvalue-basictype-member)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.calcerrorcellvalue#excel-excel-calcerrorcellvalue-basicvalue-member)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[errorSubType](/javascript/api/excel/excel.calcerrorcellvalue#excel-excel-calcerrorcellvalue-errorsubtype-member)|Represents the type of `CalcErrorCellValue`.|
||[errorType](/javascript/api/excel/excel.calcerrorcellvalue#excel-excel-calcerrorcellvalue-errortype-member)|Represents the type of `ErrorCellValue`.|
||[type](/javascript/api/excel/excel.calcerrorcellvalue#excel-excel-calcerrorcellvalue-type-member)|Represents the type of this cell value.|
|[CardLayoutListSection](/javascript/api/excel/excel.cardlayoutlistsection)|[layout](/javascript/api/excel/excel.cardlayoutlistsection#excel-excel-cardlayoutlistsection-layout-member)|Represents the type of layout for this section.|
|[CardLayoutPropertyReference](/javascript/api/excel/excel.cardlayoutpropertyreference)|[property](/javascript/api/excel/excel.cardlayoutpropertyreference#excel-excel-cardlayoutpropertyreference-property-member)|The name of the property referenced by the card layout.|
|[CardLayoutSectionStandardProperties](/javascript/api/excel/excel.cardlayoutsectionstandardproperties)|[collapsed](/javascript/api/excel/excel.cardlayoutsectionstandardproperties#excel-excel-cardlayoutsectionstandardproperties-collapsed-member)|Represents whether this section of the card is initially collapsed.|
||[collapsible](/javascript/api/excel/excel.cardlayoutsectionstandardproperties#excel-excel-cardlayoutsectionstandardproperties-collapsible-member)|Represents whether this section of the card is collapsible.|
||[properties](/javascript/api/excel/excel.cardlayoutsectionstandardproperties#excel-excel-cardlayoutsectionstandardproperties-properties-member)|Represents the names of the properties in this section.|
||[title](/javascript/api/excel/excel.cardlayoutsectionstandardproperties#excel-excel-cardlayoutsectionstandardproperties-title-member)|Represents the title of this section of the card.|
|[CardLayoutStandardProperties](/javascript/api/excel/excel.cardlayoutstandardproperties)|[mainImage](/javascript/api/excel/excel.cardlayoutstandardproperties#excel-excel-cardlayoutstandardproperties-mainimage-member)|Specifies a property which will be used as the main image of the card.|
||[sections](/javascript/api/excel/excel.cardlayoutstandardproperties#excel-excel-cardlayoutstandardproperties-sections-member)|Represents the sections of the card.|
||[subTitle](/javascript/api/excel/excel.cardlayoutstandardproperties#excel-excel-cardlayoutstandardproperties-subtitle-member)|Represents a specification of which property contains the subtitle of the card.|
||[title](/javascript/api/excel/excel.cardlayoutstandardproperties#excel-excel-cardlayoutstandardproperties-title-member)|Represents the title of the card or the specification of which property contains the title of the card.|
|[CardLayoutTableSection](/javascript/api/excel/excel.cardlayouttablesection)|[layout](/javascript/api/excel/excel.cardlayouttablesection#excel-excel-cardlayouttablesection-layout-member)|Represents the type of layout for this section.|
|[CellValueAttributionAttributes](/javascript/api/excel/excel.cellvalueattributionattributes)|[licenseAddress](/javascript/api/excel/excel.cellvalueattributionattributes#excel-excel-cellvalueattributionattributes-licenseaddress-member)|Represents a URL to a license or source that describes how this property can be used.|
||[licenseText](/javascript/api/excel/excel.cellvalueattributionattributes#excel-excel-cellvalueattributionattributes-licensetext-member)|Represents a name for the license that governs this property.|
||[sourceAddress](/javascript/api/excel/excel.cellvalueattributionattributes#excel-excel-cellvalueattributionattributes-sourceaddress-member)|Represents a URL to the source of the `CellValue`.|
||[sourceText](/javascript/api/excel/excel.cellvalueattributionattributes#excel-excel-cellvalueattributionattributes-sourcetext-member)|Represents a name for the source of the `CellValue`.|
|[CellValuePropertyMetadata](/javascript/api/excel/excel.cellvaluepropertymetadata)|[attribution](/javascript/api/excel/excel.cellvaluepropertymetadata#excel-excel-cellvaluepropertymetadata-attribution-member)|Represents attribution information to describe the source and license requirements for using this property.|
||[excludeFrom](/javascript/api/excel/excel.cellvaluepropertymetadata#excel-excel-cellvaluepropertymetadata-excludefrom-member)|Represents which features this property is excluded from.|
||[sublabel](/javascript/api/excel/excel.cellvaluepropertymetadata#excel-excel-cellvaluepropertymetadata-sublabel-member)|Represents the sublabel for this property shown in card view.|
|[CellValuePropertyMetadataExclusions](/javascript/api/excel/excel.cellvaluepropertymetadataexclusions)|[autoComplete](/javascript/api/excel/excel.cellvaluepropertymetadataexclusions#excel-excel-cellvaluepropertymetadataexclusions-autocomplete-member)|True represents that the property is excluded from the properties shown by auto complete.|
||[calcCompare](/javascript/api/excel/excel.cellvaluepropertymetadataexclusions#excel-excel-cellvaluepropertymetadataexclusions-calccompare-member)|True represents that the property is excluded from the properties used to compare cell values during recalc.|
||[cardView](/javascript/api/excel/excel.cellvaluepropertymetadataexclusions#excel-excel-cellvaluepropertymetadataexclusions-cardview-member)|True represents that the property is excluded from the properties shown by card view.|
||[dotNotation](/javascript/api/excel/excel.cellvaluepropertymetadataexclusions#excel-excel-cellvaluepropertymetadataexclusions-dotnotation-member)|True represents that the property is excluded from the properties which can be accessed via the FIELDVALUE function.|
|[CellValueProviderAttributes](/javascript/api/excel/excel.cellvalueproviderattributes)|[description](/javascript/api/excel/excel.cellvalueproviderattributes#excel-excel-cellvalueproviderattributes-description-member)|Represents the provider description property that is used in card view if no logo is specified.|
||[logoSourceAddress](/javascript/api/excel/excel.cellvalueproviderattributes#excel-excel-cellvalueproviderattributes-logosourceaddress-member)|Represents a URL used to download an image that will be used as a logo in card view.|
||[logoTargetAddress](/javascript/api/excel/excel.cellvalueproviderattributes#excel-excel-cellvalueproviderattributes-logotargetaddress-member)|Represents a URL that is the navigation target if the user clicks on the logo element in card view.|
|[Comment](/javascript/api/excel/excel.comment)|[assignTask(assignee: Identity)](/javascript/api/excel/excel.comment#excel-excel-comment-assigntask-member(1))|Assigns the task attached to the comment to the given user as an assignee.|
||[getTask()](/javascript/api/excel/excel.comment#excel-excel-comment-gettask-member(1))|Gets the task associated with this comment.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.comment#excel-excel-comment-gettaskornullobject-member(1))|Gets the task associated with this comment.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[assignTask(assignee: Identity)](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-assigntask-member(1))|Assigns the task attached to the comment to the given user as the sole assignee.|
||[getTask()](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-gettask-member(1))|Gets the task associated with this comment reply's thread.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-gettaskornullobject-member(1))|Gets the task associated with this comment reply's thread.|
|[ConnectErrorCellValue](/javascript/api/excel/excel.connecterrorcellvalue)|[basicType](/javascript/api/excel/excel.connecterrorcellvalue#excel-excel-connecterrorcellvalue-basictype-member)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.connecterrorcellvalue#excel-excel-connecterrorcellvalue-basicvalue-member)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[errorSubType](/javascript/api/excel/excel.connecterrorcellvalue#excel-excel-connecterrorcellvalue-errorsubtype-member)|Represents the type of `ConnectErrorCellValue`.|
||[errorType](/javascript/api/excel/excel.connecterrorcellvalue#excel-excel-connecterrorcellvalue-errortype-member)|Represents the type of `ErrorCellValue`.|
||[type](/javascript/api/excel/excel.connecterrorcellvalue#excel-excel-connecterrorcellvalue-type-member)|Represents the type of this cell value.|
|[Div0ErrorCellValue](/javascript/api/excel/excel.div0errorcellvalue)|[basicType](/javascript/api/excel/excel.div0errorcellvalue#excel-excel-div0errorcellvalue-basictype-member)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.div0errorcellvalue#excel-excel-div0errorcellvalue-basicvalue-member)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[errorType](/javascript/api/excel/excel.div0errorcellvalue#excel-excel-div0errorcellvalue-errortype-member)|Represents the type of `ErrorCellValue`.|
||[type](/javascript/api/excel/excel.div0errorcellvalue#excel-excel-div0errorcellvalue-type-member)|Represents the type of this cell value.|
|[DocumentTask](/javascript/api/excel/excel.documenttask)|[assignees](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-assignees-member)|Returns a collection of assignees of the task.|
||[changes](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-changes-member)|Gets the change records of the task.|
||[comment](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-comment-member)|Gets the comment associated with the task.|
||[completedBy](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-completedby-member)|Gets the most recent user to have completed the task.|
||[completedDateTime](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-completeddatetime-member)|Gets the date and time that the task was completed.|
||[createdBy](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-createdby-member)|Gets the user who created the task.|
||[createdDateTime](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-createddatetime-member)|Gets the date and time that the task was created.|
||[id](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-id-member)|Gets the ID of the task.|
||[percentComplete](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-percentcomplete-member)|Specifies the completion percentage of the task.|
||[priority](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-priority-member)|Specifies the priority of the task.|
||[setStartAndDueDateTime(startDateTime: Date, dueDateTime: Date)](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-setstartandduedatetime-member(1))|Changes the start and the due dates of the task.|
||[startAndDueDateTime](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-startandduedatetime-member)|Gets or sets the date and time the task should start and is due.|
||[title](/javascript/api/excel/excel.documenttask#excel-excel-documenttask-title-member)|Specifies title of the task.|
|[DocumentTaskChange](/javascript/api/excel/excel.documenttaskchange)|[assignee](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-assignee-member)|Represents the user assigned to the task for an `assign` change record type, or the user unassigned from the task for an `unassign` change record type.|
||[changedBy](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-changedby-member)|Represents the user who created or changed the task.|
||[commentId](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-commentid-member)|Represents the ID of the `Comment` or `CommentReply` to which the task change is anchored.|
||[createdDateTime](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-createddatetime-member)|Represents the creation date and time of the task change record.|
||[dueDateTime](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-duedatetime-member)|Represents the task's due date and time, in UTC time zone.|
||[id](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-id-member)|ID for the task change record.|
||[percentComplete](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-percentcomplete-member)|Represents the task's completion percentage.|
||[priority](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-priority-member)|Represents the task's priority.|
||[startDateTime](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-startdatetime-member)|Represents the task's start date and time, in UTC time zone.|
||[title](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-title-member)|Represents the task's title.|
||[type](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-type-member)|Represents the action type of the task change record.|
||[undoHistoryId](/javascript/api/excel/excel.documenttaskchange#excel-excel-documenttaskchange-undohistoryid-member)|Represents the `DocumentTaskChange.id` property that was undone for the `undo` change record type.|
|[DocumentTaskChangeCollection](/javascript/api/excel/excel.documenttaskchangecollection)|[getCount()](/javascript/api/excel/excel.documenttaskchangecollection#excel-excel-documenttaskchangecollection-getcount-member(1))|Gets the number of change records in the collection for the task.|
||[getItemAt(index: number)](/javascript/api/excel/excel.documenttaskchangecollection#excel-excel-documenttaskchangecollection-getitemat-member(1))|Gets a task change record by using its index in the collection.|
||[items](/javascript/api/excel/excel.documenttaskchangecollection#excel-excel-documenttaskchangecollection-items-member)|Gets the loaded child items in this collection.|
|[DocumentTaskCollection](/javascript/api/excel/excel.documenttaskcollection)|[getCount()](/javascript/api/excel/excel.documenttaskcollection#excel-excel-documenttaskcollection-getcount-member(1))|Gets the number of tasks in the collection.|
||[getItem(key: string)](/javascript/api/excel/excel.documenttaskcollection#excel-excel-documenttaskcollection-getitem-member(1))|Gets a task using its ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.documenttaskcollection#excel-excel-documenttaskcollection-getitemat-member(1))|Gets a task by its index in the collection.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.documenttaskcollection#excel-excel-documenttaskcollection-getitemornullobject-member(1))|Gets a task using its ID.|
||[items](/javascript/api/excel/excel.documenttaskcollection#excel-excel-documenttaskcollection-items-member)|Gets the loaded child items in this collection.|
|[DocumentTaskSchedule](/javascript/api/excel/excel.documenttaskschedule)|[dueDateTime](/javascript/api/excel/excel.documenttaskschedule#excel-excel-documenttaskschedule-duedatetime-member)|Gets the date and time that the task is due.|
||[startDateTime](/javascript/api/excel/excel.documenttaskschedule#excel-excel-documenttaskschedule-startdatetime-member)|Gets the date and time that the task should start.|
|[DoubleCellValue](/javascript/api/excel/excel.doublecellvalue)|[basicType](/javascript/api/excel/excel.doublecellvalue#excel-excel-doublecellvalue-basictype-member)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.doublecellvalue#excel-excel-doublecellvalue-basicvalue-member)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[type](/javascript/api/excel/excel.doublecellvalue#excel-excel-doublecellvalue-type-member)|Represents the type of this cell value.|
|[EmptyCellValue](/javascript/api/excel/excel.emptycellvalue)|[basicType](/javascript/api/excel/excel.emptycellvalue#excel-excel-emptycellvalue-basictype-member)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.emptycellvalue#excel-excel-emptycellvalue-basicvalue-member)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[type](/javascript/api/excel/excel.emptycellvalue#excel-excel-emptycellvalue-type-member)|Represents the type of this cell value.|
|[EntityCardLayout](/javascript/api/excel/excel.entitycardlayout)|[layout](/javascript/api/excel/excel.entitycardlayout#excel-excel-entitycardlayout-layout-member)|Represent the type of this layout.|
|[EntityCellValue](/javascript/api/excel/excel.entitycellvalue)|[basicType](/javascript/api/excel/excel.entitycellvalue#excel-excel-entitycellvalue-basictype-member)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.entitycellvalue#excel-excel-entitycellvalue-basicvalue-member)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[cardLayout](/javascript/api/excel/excel.entitycellvalue#excel-excel-entitycellvalue-cardlayout-member)|Represents the layout of this entity in card view.|
||[properties: {            [key: string]](/javascript/api/excel/excel.entitycellvalue#excel-excel-entitycellvalue-properties-member)|Represents the properties of this entity and their metadata.|
||[text](/javascript/api/excel/excel.entitycellvalue#excel-excel-entitycellvalue-text-member)|Represents the text shown when a cell with this value is rendered.|
||[type](/javascript/api/excel/excel.entitycellvalue#excel-excel-entitycellvalue-type-member)|Represents the type of this cell value.|
|[FieldErrorCellValue](/javascript/api/excel/excel.fielderrorcellvalue)|[basicType](/javascript/api/excel/excel.fielderrorcellvalue#excel-excel-fielderrorcellvalue-basictype-member)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.fielderrorcellvalue#excel-excel-fielderrorcellvalue-basicvalue-member)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[errorSubType](/javascript/api/excel/excel.fielderrorcellvalue#excel-excel-fielderrorcellvalue-errorsubtype-member)|Represents the type of `FieldErrorCellValue`.|
||[errorType](/javascript/api/excel/excel.fielderrorcellvalue#excel-excel-fielderrorcellvalue-errortype-member)|Represents the type of `ErrorCellValue`.|
||[type](/javascript/api/excel/excel.fielderrorcellvalue#excel-excel-fielderrorcellvalue-type-member)|Represents the type of this cell value.|
|[FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue)|[basicType](/javascript/api/excel/excel.formattednumbercellvalue#excel-excel-formattednumbercellvalue-basictype-member)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.formattednumbercellvalue#excel-excel-formattednumbercellvalue-basicvalue-member)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[numberFormat](/javascript/api/excel/excel.formattednumbercellvalue#excel-excel-formattednumbercellvalue-numberformat-member)|Returns the number format string that is used to display this value.|
||[type](/javascript/api/excel/excel.formattednumbercellvalue#excel-excel-formattednumbercellvalue-type-member)|Represents the type of this cell value.|
|[GettingDataErrorCellValue](/javascript/api/excel/excel.gettingdataerrorcellvalue)|[basicType](/javascript/api/excel/excel.gettingdataerrorcellvalue#excel-excel-gettingdataerrorcellvalue-basictype-member)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.gettingdataerrorcellvalue#excel-excel-gettingdataerrorcellvalue-basicvalue-member)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[errorType](/javascript/api/excel/excel.gettingdataerrorcellvalue#excel-excel-gettingdataerrorcellvalue-errortype-member)|Represents the type of `ErrorCellValue`.|
||[type](/javascript/api/excel/excel.gettingdataerrorcellvalue#excel-excel-gettingdataerrorcellvalue-type-member)|Represents the type of this cell value.|
|[Identity](/javascript/api/excel/excel.identity)|[displayName](/javascript/api/excel/excel.identity#excel-excel-identity-displayname-member)|Represents the user's display name.|
||[email](/javascript/api/excel/excel.identity#excel-excel-identity-email-member)|Represents the user's email address.|
||[id](/javascript/api/excel/excel.identity#excel-excel-identity-id-member)|Represents the user's unique ID.|
|[IdentityCollection](/javascript/api/excel/excel.identitycollection)|[add(assignee: Identity)](/javascript/api/excel/excel.identitycollection#excel-excel-identitycollection-add-member(1))|Adds a user identity to the collection.|
||[clear()](/javascript/api/excel/excel.identitycollection#excel-excel-identitycollection-clear-member(1))|Removes all user identities from the collection.|
||[getCount()](/javascript/api/excel/excel.identitycollection#excel-excel-identitycollection-getcount-member(1))|Gets the number of items in the collection.|
||[getItemAt(index: number)](/javascript/api/excel/excel.identitycollection#excel-excel-identitycollection-getitemat-member(1))|Gets a document user identity by using its index in the collection.|
||[items](/javascript/api/excel/excel.identitycollection#excel-excel-identitycollection-items-member)|Gets the loaded child items in this collection.|
||[remove(assignee: Identity)](/javascript/api/excel/excel.identitycollection#excel-excel-identitycollection-remove-member(1))|Removes a user identity from the collection.|
|[IdentityEntity](/javascript/api/excel/excel.identityentity)|[displayName](/javascript/api/excel/excel.identityentity#excel-excel-identityentity-displayname-member)|Represents the user's display name.|
||[email](/javascript/api/excel/excel.identityentity#excel-excel-identityentity-email-member)|Represents the user's email address.|
||[id](/javascript/api/excel/excel.identityentity#excel-excel-identityentity-id-member)|Represents the user's unique ID.|
|[LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|[dataProvider](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-dataprovider-member)|The name of the data provider for the linked data type.|
||[lastRefreshed](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-lastrefreshed-member)|The local time-zone date and time since the workbook was opened when the linked data type was last refreshed.|
||[name](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-name-member)|The name of the linked data type.|
||[periodicRefreshInterval](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-periodicrefreshinterval-member)|The frequency, in seconds, at which the linked data type is refreshed if `refreshMode` is set to "Periodic".|
||[refreshMode](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-refreshmode-member)|The mechanism by which the data for the linked data type is retrieved.|
||[requestRefresh()](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-requestrefresh-member(1))|Makes a request to refresh the linked data type.|
||[requestSetRefreshMode(refreshMode: Excel.LinkedDataTypeRefreshMode)](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-requestsetrefreshmode-member(1))|Makes a request to change the refresh mode for this linked data type.|
||[serviceId](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-serviceid-member)|The unique ID of the linked data type.|
||[supportedRefreshModes](/javascript/api/excel/excel.linkeddatatype#excel-excel-linkeddatatype-supportedrefreshmodes-member)|Returns an array with all the refresh modes supported by the linked data type.|
|[LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs)|[serviceId](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#excel-excel-linkeddatatypeaddedeventargs-serviceid-member)|The unique ID of the new linked data type.|
||[source](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#excel-excel-linkeddatatypeaddedeventargs-source-member)|Gets the source of the event.|
||[type](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#excel-excel-linkeddatatypeaddedeventargs-type-member)|Gets the type of the event.|
|[LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection)|[getCount()](/javascript/api/excel/excel.linkeddatatypecollection#excel-excel-linkeddatatypecollection-getcount-member(1))|Gets the number of linked data types in the collection.|
||[getItem(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#excel-excel-linkeddatatypecollection-getitem-member(1))|Gets a linked data type by service ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.linkeddatatypecollection#excel-excel-linkeddatatypecollection-getitemat-member(1))|Gets a linked data type by its index in the collection.|
||[getItemOrNullObject(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#excel-excel-linkeddatatypecollection-getitemornullobject-member(1))|Gets a linked data type by ID.|
||[items](/javascript/api/excel/excel.linkeddatatypecollection#excel-excel-linkeddatatypecollection-items-member)|Gets the loaded child items in this collection.|
||[requestRefreshAll()](/javascript/api/excel/excel.linkeddatatypecollection#excel-excel-linkeddatatypecollection-requestrefreshall-member(1))|Makes a request to refresh all the linked data types in the collection.|
|[LinkedEntityCellValue](/javascript/api/excel/excel.linkedentitycellvalue)|[basicType](/javascript/api/excel/excel.linkedentitycellvalue#excel-excel-linkedentitycellvalue-basictype-member)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.linkedentitycellvalue#excel-excel-linkedentitycellvalue-basicvalue-member)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[id](/javascript/api/excel/excel.linkedentitycellvalue#excel-excel-linkedentitycellvalue-id-member)|Represents the service source that provided the information in this value.|
||[properties: {            [key: string]: CellValue & {                propertyMetadata](/javascript/api/excel/excel.linkedentitycellvalue#excel-excel-linkedentitycellvalue-properties-member)|Represents the properties of this entity and their metadata.|
||[propertyMetadata](/javascript/api/excel/excel.linkedentitycellvalue#excel-excel-linkedentitycellvalue-propertymetadata-member)||
||[provider](/javascript/api/excel/excel.linkedentitycellvalue#excel-excel-linkedentitycellvalue-provider-member)|Represents information that describes the service which provided the image.|
||[text](/javascript/api/excel/excel.linkedentitycellvalue#excel-excel-linkedentitycellvalue-text-member)|Represents the text shown when a cell with this value is rendered.|
||[type](/javascript/api/excel/excel.linkedentitycellvalue#excel-excel-linkedentitycellvalue-type-member)|Represents the type of this cell value.|
|[LinkedEntityId](/javascript/api/excel/excel.linkedentityid)|[culture](/javascript/api/excel/excel.linkedentityid#excel-excel-linkedentityid-culture-member)|Represents which language culture was used to create this `CellValue`.|
||[domainId](/javascript/api/excel/excel.linkedentityid#excel-excel-linkedentityid-domainid-member)|Represents a domain specific to a service used to create the `CellValue`.|
||[entityId](/javascript/api/excel/excel.linkedentityid#excel-excel-linkedentityid-entityid-member)|Represents an identifier specific to a service used to create the `CellValue`.|
||[serviceId](/javascript/api/excel/excel.linkedentityid#excel-excel-linkedentityid-serviceid-member)|Represents which service was used to create the `CellValue`.|
|[NameErrorCellValue](/javascript/api/excel/excel.nameerrorcellvalue)|[basicType](/javascript/api/excel/excel.nameerrorcellvalue#excel-excel-nameerrorcellvalue-basictype-member)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.nameerrorcellvalue#excel-excel-nameerrorcellvalue-basicvalue-member)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[errorType](/javascript/api/excel/excel.nameerrorcellvalue#excel-excel-nameerrorcellvalue-errortype-member)|Represents the type of `ErrorCellValue`.|
||[type](/javascript/api/excel/excel.nameerrorcellvalue#excel-excel-nameerrorcellvalue-type-member)|Represents the type of this cell value.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[valueAsJson](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-valueasjson-member)|A JSON representation of the values in this named item.|
|[NamedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|[valuesAsJson](/javascript/api/excel/excel.nameditemarrayvalues#excel-excel-nameditemarrayvalues-valuesasjson-member)|A JSON representation of the values in the cells in this range.|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-getitemornullobject-member(1))|Gets a sheet view using its name.|
|[NotAvailableErrorCellValue](/javascript/api/excel/excel.notavailableerrorcellvalue)|[basicType](/javascript/api/excel/excel.notavailableerrorcellvalue#excel-excel-notavailableerrorcellvalue-basictype-member)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.notavailableerrorcellvalue#excel-excel-notavailableerrorcellvalue-basicvalue-member)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[errorType](/javascript/api/excel/excel.notavailableerrorcellvalue#excel-excel-notavailableerrorcellvalue-errortype-member)|Represents the type of `ErrorCellValue`.|
||[type](/javascript/api/excel/excel.notavailableerrorcellvalue#excel-excel-notavailableerrorcellvalue-type-member)|Represents the type of this cell value.|
|[NullErrorCellValue](/javascript/api/excel/excel.nullerrorcellvalue)|[basicType](/javascript/api/excel/excel.nullerrorcellvalue#excel-excel-nullerrorcellvalue-basictype-member)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.nullerrorcellvalue#excel-excel-nullerrorcellvalue-basicvalue-member)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[errorType](/javascript/api/excel/excel.nullerrorcellvalue#excel-excel-nullerrorcellvalue-errortype-member)|Represents the type of `ErrorCellValue`.|
||[type](/javascript/api/excel/excel.nullerrorcellvalue#excel-excel-nullerrorcellvalue-type-member)|Represents the type of this cell value.|
|[NumErrorCellValue](/javascript/api/excel/excel.numerrorcellvalue)|[basicType](/javascript/api/excel/excel.numerrorcellvalue#excel-excel-numerrorcellvalue-basictype-member)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.numerrorcellvalue#excel-excel-numerrorcellvalue-basicvalue-member)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[errorType](/javascript/api/excel/excel.numerrorcellvalue#excel-excel-numerrorcellvalue-errortype-member)|Represents the type of `ErrorCellValue`.|
||[type](/javascript/api/excel/excel.numerrorcellvalue#excel-excel-numerrorcellvalue-type-member)|Represents the type of this cell value.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-getcell-member(1))|Gets a unique cell in the PivotTable based on a data hierarchy and the row and column items of their respective hierarchies.|
||[pivotStyle](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-pivotstyle-member)|The style applied to the PivotTable.|
||[setStyle(style: string \| PivotTableStyle \| BuiltInPivotTableStyle)](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-setstyle-member(1))|Sets the style applied to the PivotTable.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[getDataSourceString()](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-getdatasourcestring-member(1))|Returns the string representation of the data source for the PivotTable.|
||[getDataSourceType()](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-getdatasourcetype-member(1))|Gets the type of the data source for the PivotTable.|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getFirstOrNullObject()](/javascript/api/excel/excel.pivottablescopedcollection#excel-excel-pivottablescopedcollection-getfirstornullobject-member(1))|Gets the first PivotTable in the collection.|
|[PlaceholderErrorCellValue](/javascript/api/excel/excel.placeholdererrorcellvalue)|[basicType](/javascript/api/excel/excel.placeholdererrorcellvalue#excel-excel-placeholdererrorcellvalue-basictype-member)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.placeholdererrorcellvalue#excel-excel-placeholdererrorcellvalue-basicvalue-member)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[errorType](/javascript/api/excel/excel.placeholdererrorcellvalue#excel-excel-placeholdererrorcellvalue-errortype-member)|Represents the type of `ErrorCellValue`.|
||[target](/javascript/api/excel/excel.placeholdererrorcellvalue#excel-excel-placeholdererrorcellvalue-target-member)|`PlaceholderErrorCellValue` is used during processing, while data is downloaded.|
||[type](/javascript/api/excel/excel.placeholdererrorcellvalue#excel-excel-placeholdererrorcellvalue-type-member)|Represents the type of this cell value.|
|[Range](/javascript/api/excel/excel.range)|[getDependents()](/javascript/api/excel/excel.range#excel-excel-range-getdependents-member(1))|Returns a `WorkbookRangeAreas` object that represents the range containing all the dependents of a cell in the same worksheet or in multiple worksheets.|
||[valuesAsJson](/javascript/api/excel/excel.range#excel-excel-range-valuesasjson-member)|A JSON representation of the values in the cells in this range.|
|[RangeView](/javascript/api/excel/excel.rangeview)|[valuesAsJson](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-valuesasjson-member)|A JSON representation of the values in the cells in this range.|
|[RefErrorCellValue](/javascript/api/excel/excel.referrorcellvalue)|[basicType](/javascript/api/excel/excel.referrorcellvalue#excel-excel-referrorcellvalue-basictype-member)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.referrorcellvalue#excel-excel-referrorcellvalue-basicvalue-member)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[errorSubType](/javascript/api/excel/excel.referrorcellvalue#excel-excel-referrorcellvalue-errorsubtype-member)|Represents the type of `RefErrorCellValue`.|
||[errorType](/javascript/api/excel/excel.referrorcellvalue#excel-excel-referrorcellvalue-errortype-member)|Represents the type of `ErrorCellValue`.|
||[type](/javascript/api/excel/excel.referrorcellvalue#excel-excel-referrorcellvalue-type-member)|Represents the type of this cell value.|
|[RefreshModeChangedEventArgs](/javascript/api/excel/excel.refreshmodechangedeventargs)|[refreshMode](/javascript/api/excel/excel.refreshmodechangedeventargs#excel-excel-refreshmodechangedeventargs-refreshmode-member)|The linked data type refresh mode.|
||[serviceId](/javascript/api/excel/excel.refreshmodechangedeventargs#excel-excel-refreshmodechangedeventargs-serviceid-member)|The unique ID of the object whose refresh mode was changed.|
||[source](/javascript/api/excel/excel.refreshmodechangedeventargs#excel-excel-refreshmodechangedeventargs-source-member)|Gets the source of the event.|
||[type](/javascript/api/excel/excel.refreshmodechangedeventargs#excel-excel-refreshmodechangedeventargs-type-member)|Gets the type of the event.|
|[RefreshRequestCompletedEventArgs](/javascript/api/excel/excel.refreshrequestcompletedeventargs)|[refreshed](/javascript/api/excel/excel.refreshrequestcompletedeventargs#excel-excel-refreshrequestcompletedeventargs-refreshed-member)|Indicates if the request to refresh was successful.|
||[serviceId](/javascript/api/excel/excel.refreshrequestcompletedeventargs#excel-excel-refreshrequestcompletedeventargs-serviceid-member)|The unique ID of the object whose refresh request was completed.|
||[source](/javascript/api/excel/excel.refreshrequestcompletedeventargs#excel-excel-refreshrequestcompletedeventargs-source-member)|Gets the source of the event.|
||[type](/javascript/api/excel/excel.refreshrequestcompletedeventargs#excel-excel-refreshrequestcompletedeventargs-type-member)|Gets the type of the event.|
||[warnings](/javascript/api/excel/excel.refreshrequestcompletedeventargs#excel-excel-refreshrequestcompletedeventargs-warnings-member)|An array that contains any warnings generated from the refresh request.|
|[Shape](/javascript/api/excel/excel.shape)|[displayName](/javascript/api/excel/excel.shape#excel-excel-shape-displayname-member)|Gets the display name of the shape.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addsvg-member(1))|Creates a scalable vector graphic (SVG) from an XML string and adds it to the worksheet.|
|[Slicer](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#excel-excel-slicer-nameinformula-member)|Represents the slicer name used in the formula.|
||[setStyle(style: string \| SlicerStyle \| BuiltInSlicerStyle)](/javascript/api/excel/excel.slicer#excel-excel-slicer-setstyle-member(1))|Sets the style applied to the slicer.|
||[slicerStyle](/javascript/api/excel/excel.slicer#excel-excel-slicer-slicerstyle-member)|The style applied to the slicer.|
|[SpillErrorCellValue](/javascript/api/excel/excel.spillerrorcellvalue)|[basicType](/javascript/api/excel/excel.spillerrorcellvalue#excel-excel-spillerrorcellvalue-basictype-member)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.spillerrorcellvalue#excel-excel-spillerrorcellvalue-basicvalue-member)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[errorSubType](/javascript/api/excel/excel.spillerrorcellvalue#excel-excel-spillerrorcellvalue-errorsubtype-member)|Represents the type of `SpillErrorCellValue`.|
||[errorType](/javascript/api/excel/excel.spillerrorcellvalue#excel-excel-spillerrorcellvalue-errortype-member)|Represents the type of `ErrorCellValue`.|
||[spilledColumns](/javascript/api/excel/excel.spillerrorcellvalue#excel-excel-spillerrorcellvalue-spilledcolumns-member)|Represents the number of columns that would spill if there were no #SPILL! error.|
||[spilledRows](/javascript/api/excel/excel.spillerrorcellvalue#excel-excel-spillerrorcellvalue-spilledrows-member)|Represents the number of rows that would spill if there were no #SPILL! error.|
||[type](/javascript/api/excel/excel.spillerrorcellvalue#excel-excel-spillerrorcellvalue-type-member)|Represents the type of this cell value.|
|[StringCellValue](/javascript/api/excel/excel.stringcellvalue)|[basicType](/javascript/api/excel/excel.stringcellvalue#excel-excel-stringcellvalue-basictype-member)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.stringcellvalue#excel-excel-stringcellvalue-basicvalue-member)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[type](/javascript/api/excel/excel.stringcellvalue#excel-excel-stringcellvalue-type-member)|Represents the type of this cell value.|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#excel-excel-table-clearstyle-member(1))|Changes the table to use the default table style.|
||[onFiltered](/javascript/api/excel/excel.table#excel-excel-table-onfiltered-member)|Occurs when a filter is applied on a specific table.|
||[setStyle(style: string \| TableStyle \| BuiltInTableStyle)](/javascript/api/excel/excel.table#excel-excel-table-setstyle-member(1))|Sets the style applied to the table.|
||[tableStyle](/javascript/api/excel/excel.table#excel-excel-table-tablestyle-member)|The style applied to the table.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-onfiltered-member)|Occurs when a filter is applied on any table in a workbook, or a worksheet.|
|[TableColumn](/javascript/api/excel/excel.tablecolumn)|[valuesAsJson](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-valuesasjson-member)|A JSON representation of the values in the cells in this table column.|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#excel-excel-tablefilteredeventargs-tableid-member)|Gets the ID of the table in which the filter is applied.|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#excel-excel-tablefilteredeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#excel-excel-tablefilteredeventargs-worksheetid-member)|Gets the ID of the worksheet which contains the table.|
|[TableRow](/javascript/api/excel/excel.tablerow)|[valuesAsJson](/javascript/api/excel/excel.tablerow#excel-excel-tablerow-valuesasjson-member)|A JSON representation of the values in the cells in this table row.|
|[ValueErrorCellValue](/javascript/api/excel/excel.valueerrorcellvalue)|[basicType](/javascript/api/excel/excel.valueerrorcellvalue#excel-excel-valueerrorcellvalue-basictype-member)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.valueerrorcellvalue#excel-excel-valueerrorcellvalue-basicvalue-member)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[errorSubType](/javascript/api/excel/excel.valueerrorcellvalue#excel-excel-valueerrorcellvalue-errorsubtype-member)|Represents the type of `ValueErrorCellValue`.|
||[errorType](/javascript/api/excel/excel.valueerrorcellvalue#excel-excel-valueerrorcellvalue-errortype-member)|Represents the type of `ErrorCellValue`.|
||[type](/javascript/api/excel/excel.valueerrorcellvalue#excel-excel-valueerrorcellvalue-type-member)|Represents the type of this cell value.|
|[ValueTypeNotAvailableCellValue](/javascript/api/excel/excel.valuetypenotavailablecellvalue)|[basicType](/javascript/api/excel/excel.valuetypenotavailablecellvalue#excel-excel-valuetypenotavailablecellvalue-basictype-member)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.valuetypenotavailablecellvalue#excel-excel-valuetypenotavailablecellvalue-basicvalue-member)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[type](/javascript/api/excel/excel.valuetypenotavailablecellvalue#excel-excel-valuetypenotavailablecellvalue-type-member)|Represents the type of this cell value.|
|[WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue)|[address](/javascript/api/excel/excel.webimagecellvalue#excel-excel-webimagecellvalue-address-member)|Represents the URL from which the image will be downloaded.|
||[altText](/javascript/api/excel/excel.webimagecellvalue#excel-excel-webimagecellvalue-alttext-member)|Represents the alternate text that can be used in accessibility scenarios to describe what the image represents.|
||[attribution](/javascript/api/excel/excel.webimagecellvalue#excel-excel-webimagecellvalue-attribution-member)|Represents attribution information to describe the source and license requirements for using this image.|
||[basicType](/javascript/api/excel/excel.webimagecellvalue#excel-excel-webimagecellvalue-basictype-member)|Represents the value that would be returned by `Range.valueTypes` for a cell with this value.|
||[basicValue](/javascript/api/excel/excel.webimagecellvalue#excel-excel-webimagecellvalue-basicvalue-member)|Represents the value that would be returned by `Range.values` for a cell with this value.|
||[provider](/javascript/api/excel/excel.webimagecellvalue#excel-excel-webimagecellvalue-provider-member)|Represents information that describes the entity or individual who provided the image.|
||[relatedImagesAddress](/javascript/api/excel/excel.webimagecellvalue#excel-excel-webimagecellvalue-relatedimagesaddress-member)|Represents the URL of a webpage with images that are considered related to this `WebImageCellValue`.|
||[type](/javascript/api/excel/excel.webimagecellvalue#excel-excel-webimagecellvalue-type-member)|Represents the type of this cell value.|
|[Workbook](/javascript/api/excel/excel.workbook)|[getLinkedEntityCellValue(linkedEntityCellValueId: LinkedEntityId)](/javascript/api/excel/excel.workbook#excel-excel-workbook-getlinkedentitycellvalue-member(1))|Returns a `LinkedEntityCellValue` based on the provided `LinkedEntityId`.|
||[linkedDataTypes](/javascript/api/excel/excel.workbook#excel-excel-workbook-linkeddatatypes-member)|Returns a collection of linked data types that are part of the workbook.|
||[showPivotFieldList](/javascript/api/excel/excel.workbook#excel-excel-workbook-showpivotfieldlist-member)|Specifies whether the PivotTable's field list pane is shown at the workbook level.|
||[tasks](/javascript/api/excel/excel.workbook#excel-excel-workbook-tasks-member)|Returns a collection of tasks that are present in the workbook.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#excel-excel-workbook-use1904datesystem-member)|True if the workbook uses the 1904 date system.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onFiltered](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onfiltered-member)|Occurs when a filter is applied on a specific worksheet.|
||[tasks](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-tasks-member)|Returns a collection of tasks that are present in the worksheet.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-addfrombase64-member(1))|Inserts the specified worksheets of a workbook into the current workbook.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onfiltered-member)|Occurs when any worksheet's filter is applied in the workbook.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#excel-excel-worksheetfilteredeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#excel-excel-worksheetfilteredeventargs-worksheetid-member)|Gets the ID of the worksheet in which the filter is applied.|
|[WorksheetProtection](/javascript/api/excel/excel.worksheetprotection)|[allowEditRanges](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-alloweditranges-member)|Specifies the `AllowEditRangeCollection` object found in this worksheet.|
||[canPauseProtection](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-canpauseprotection-member)|Specifies if protection can be paused for this worksheet.|
||[checkPassword(password?: string)](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-checkpassword-member(1))|Specifies if the password can be used to unlock worksheet protection.|
||[isPasswordProtected](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-ispasswordprotected-member)|Specifies if the sheet is password protected.|
||[isPaused](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-ispaused-member)|Specifies if worksheet protection is paused.|
||[pauseProtection(password?: string)](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-pauseprotection-member(1))|Pauses worksheet protection for the given worksheet object for the user in a given session.|
||[resumeProtection()](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-resumeprotection-member(1))|Resumes worksheet protection for the given worksheet object for the user in a given session.|
||[setPassword(password?: string)](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-setpassword-member(1))|Changes the password associated with the `WorksheetProtection` object.|
||[updateOptions(options: Excel.WorksheetProtectionOptions)](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-updateoptions-member(1))|Change the worksheet protection options associated to the `WorksheetProtection` object.|
|[WorksheetProtectionChangedEventArgs](/javascript/api/excel/excel.worksheetprotectionchangedeventargs)|[allowEditRangesChanged](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#excel-excel-worksheetprotectionchangedeventargs-alloweditrangeschanged-member)|Specifies if any of the `AllowEditRange` objects have changed.|
||[protectionOptionsChanged](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#excel-excel-worksheetprotectionchangedeventargs-protectionoptionschanged-member)|Specifies if the `WorksheetProtectionOptions` have changed.|
||[sheetPasswordChanged](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#excel-excel-worksheetprotectionchangedeventargs-sheetpasswordchanged-member)|Specifies if the worksheet password has changed.|

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel?view=excel-js-preview&preserve-view=true)
- [Excel JavaScript API requirement sets](excel-api-requirement-sets.md)
