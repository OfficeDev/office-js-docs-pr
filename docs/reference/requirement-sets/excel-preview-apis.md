---
title: Excel JavaScript preview APIs
description: 'Details about upcoming Excel JavaScript APIs'
ms.date: 09/24/2019
ms.prod: excel
localization_priority: Normal
---

# Excel JavaScript preview APIs

New Excel JavaScript APIs are first introduced in "preview" and later become part of a specific, numbered requirement set after sufficient testing occurs and user feedback is acquired.

The first table provides a concise summary of the APIs, while the subsequent table gives a detailed list.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| Feature area | Description | Relevant objects |
|:--- |:--- |:--- |
| [Slicer](../../excel/excel-add-ins-pivottables.md#slicers-preview) | Insert and configure slicers to tables and PivotTables. | [Slicer](/javascript/api/excel/excel.slicer) |
| [Comments](../../excel/excel-add-ins-workbooks.md#comments-preview) | Add, edit, and delete comments. | [Comment](/javascript/api/excel/excel.comment), [CommentCollection](/javascript/api/excel/excel.commentcollection) |
| Workbook [Save](../../excel/excel-add-ins-workbooks.md#save-the-workbook-preview) and [Close](../../excel/excel-add-ins-workbooks.md#close-the-workbook-preview) | Save and close workbooks.  | [Workbook](/javascript/api/excel/excel.workbook) |
| [Insert Workbook](../../excel/excel-add-ins-workbooks.md#insert-a-copy-of-an-existing-workbook-into-the-current-one-preview) | Insert one workbook into another.  | [Workbook](/javascript/api/excel/excel.worksheetcollection) |

## API list

The following table lists the Excel JavaScript APIs currently in preview. To see a complete list of all Excel JavaScript APIs (including preview APIs and previously released APIs), see [all Excel JavaScript APIs](/javascript/api/excel?view=excel-js-preview).

| Class | Fields | Description |
|:---|:---|:---|
|[Comment](/javascript/api/excel/excel.comment)|[content](/javascript/api/excel/excel.comment#content)|Gets or sets the comment's content. The string is plain text.|
||[delete()](/javascript/api/excel/excel.comment#delete--)|Deletes the comment and all the connected replies.|
||[getLocation()](/javascript/api/excel/excel.comment#getlocation--)|Gets the cell where this comment is located.|
||[authorEmail](/javascript/api/excel/excel.comment#authoremail)|Gets the email of the comment's author.|
||[authorName](/javascript/api/excel/excel.comment#authorname)|Gets the name of the comment's author.|
||[creationDate](/javascript/api/excel/excel.comment#creationdate)|Gets the creation time of the comment. Returns null if the comment was converted from a note, since the comment does not have a creation date.|
||[id](/javascript/api/excel/excel.comment#id)|Represents the comment identifier. Read-only.|
||[mentions](/javascript/api/excel/excel.comment#mentions)|Gets the entities (e.g. people) that are mentioned in comments.|
||[replies](/javascript/api/excel/excel.comment#replies)|Represents a collection of reply objects associated with the comment. Read-only.|
||[richContent](/javascript/api/excel/excel.comment#richcontent)|Gets the rich comment content (e.g. mentions in comments). This string is not meant to be displayed to end-users. Your add-in should only use this to parse rich comment content.|
||[resolved](/javascript/api/excel/excel.comment#resolved)|Gets or sets the comment thread status. A value of "true" means the comment thread is in the resolved state.|
||[updateMentions(contentWithMentions: Excel.CommentRichContent)](/javascript/api/excel/excel.comment#updatementions-contentwithmentions-)|Updates the comment content with a specially formatted string and a list of mentions.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add(cellAddress: Range \| string, content: CommentRichContent \| string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentcollection#add-celladdress--content--contenttype-)|Creates a new comment with the given content on the given cell. An `InvalidArgument` error is thrown if the provided range is larger than one cell.|
||[getCount()](/javascript/api/excel/excel.commentcollection#getcount--)|Gets the number of comments in the collection.|
||[getItem(commentId: string)](/javascript/api/excel/excel.commentcollection#getitem-commentid-)|Gets a comment from the collection based on its ID. Read-only.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentcollection#getitemat-index-)|Gets a comment from the collection based on its position.|
||[getItemByCell(cellAddress: Range \| string)](/javascript/api/excel/excel.commentcollection#getitembycell-celladdress-)|Gets the comment from the specified cell.|
||[getItemByReplyId(replyId: string)](/javascript/api/excel/excel.commentcollection#getitembyreplyid-replyid-)|Gets the comment to which the given reply is connected.|
||[items](/javascript/api/excel/excel.commentcollection#items)|Gets the loaded child items in this collection.|
|[CommentMention](/javascript/api/excel/excel.commentmention)|[email](/javascript/api/excel/excel.commentmention#email)|Gets or sets the email address of the entity that is mentioned in comment.|
||[id](/javascript/api/excel/excel.commentmention#id)|Gets or sets the id of the entity. This is aligned with the id information in `CommentRichContent.richContent`.|
||[name](/javascript/api/excel/excel.commentmention#name)|Gets or sets the name of the entity that is mentioned in comment.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[content](/javascript/api/excel/excel.commentreply#content)|Gets or sets the comment reply's content. The string is plain text.|
||[delete()](/javascript/api/excel/excel.commentreply#delete--)|Deletes the comment reply.|
||[getLocation()](/javascript/api/excel/excel.commentreply#getlocation--)|Gets the cell where this comment reply is located.|
||[getParentComment()](/javascript/api/excel/excel.commentreply#getparentcomment--)|Gets the parent comment of this reply.|
||[authorEmail](/javascript/api/excel/excel.commentreply#authoremail)|Gets the email of the comment reply's author.|
||[authorName](/javascript/api/excel/excel.commentreply#authorname)|Gets the name of the comment reply's author.|
||[creationDate](/javascript/api/excel/excel.commentreply#creationdate)|Gets the creation time of the comment reply.|
||[id](/javascript/api/excel/excel.commentreply#id)|Represents the comment reply identifier. Read-only.|
||[mentions](/javascript/api/excel/excel.commentreply#mentions)|Gets the entities (e.g. people) that are mentioned in comments.|
||[resolved](/javascript/api/excel/excel.commentreply#resolved)|Gets or sets the comment reply status. A value of "true" means the comment reply is in the resolved state.|
||[richContent](/javascript/api/excel/excel.commentreply#richcontent)|Gets the rich comment content (e.g. mentions in comments). This string is not meant to be displayed to end-users. Your add-in should only use this to parse rich comment content.|
||[updateMentions(contentWithMentions: Excel.CommentRichContent)](/javascript/api/excel/excel.commentreply#updatementions-contentwithmentions-)|Updates the comment content with a specially formatted string and a list of mentions.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add(content: CommentRichContent \| string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|Creates a comment reply for comment.|
||[getCount()](/javascript/api/excel/excel.commentreplycollection#getcount--)|Gets the number of comment replies in the collection.|
||[getItem(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#getitem-commentreplyid-)|Returns a comment reply identified by its ID. Read-only.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentreplycollection#getitemat-index-)|Gets a comment reply based on its position in the collection.|
||[items](/javascript/api/excel/excel.commentreplycollection#items)|Gets the loaded child items in this collection.|
|[CommentRichContent](/javascript/api/excel/excel.commentrichcontent)|[mentions](/javascript/api/excel/excel.commentrichcontent#mentions)|An array containing all the entities (e.g. people) mentioned within the comment.|
||[richContent](/javascript/api/excel/excel.commentrichcontent#richcontent)||
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[enableFieldList](/javascript/api/excel/excel.pivotlayout#enablefieldlist)|Specifies whether the field list can be shown in the UI.|
||[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|Gets a unique cell in the PivotTable based on a data hierarchy and the row and column items of their respective hierarchies. The returned cell is the intersection of the given row and column that contains the data from the given hierarchy. This method is the inverse of calling getPivotItems and getDataHierarchy on a particular cell.|
|[PivotTableStyle](/javascript/api/excel/excel.pivottablestyle)|[delete()](/javascript/api/excel/excel.pivottablestyle#delete--)|Deletes the PivotTableStyle.|
||[duplicate()](/javascript/api/excel/excel.pivottablestyle#duplicate--)|Creates a duplicate of this PivotTableStyle with copies of all the style elements.|
||[name](/javascript/api/excel/excel.pivottablestyle#name)|Gets the name of the PivotTableStyle.|
||[readOnly](/javascript/api/excel/excel.pivottablestyle#readonly)|Specifies whether this PivotTableStyle object is read-only. Read-only.|
|[PivotTableStyleCollection](/javascript/api/excel/excel.pivottablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.pivottablestylecollection#add-name--makeuniquename-)|Creates a blank PivotTableStyle with the specified name.|
||[getCount()](/javascript/api/excel/excel.pivottablestylecollection#getcount--)|Gets the number of PivotTable styles in the collection.|
||[getDefault()](/javascript/api/excel/excel.pivottablestylecollection#getdefault--)|Gets the default PivotTableStyle for the parent object's scope.|
||[getItem(name: string)](/javascript/api/excel/excel.pivottablestylecollection#getitem-name-)|Gets a PivotTableStyle by name.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablestylecollection#getitemornullobject-name-)|Gets a PivotTableStyle by name. If the PivotTableStyle does not exist, will return a null object.|
||[items](/javascript/api/excel/excel.pivottablestylecollection#items)|Gets the loaded child items in this collection.|
||[setDefault(newDefaultStyle: PivotTableStyle \| string)](/javascript/api/excel/excel.pivottablestylecollection#setdefault-newdefaultstyle-)|Sets the default PivotTableStyle for use in the parent object's scope.|
|[Range](/javascript/api/excel/excel.range)|[getSpillParent()](/javascript/api/excel/excel.range#getspillparent--)|Gets the range object containing the anchor cell for a cell getting spilled into. Fails if applied to a range with more than one cell. Read-only.|
||[getSpillParentOrNullObject()](/javascript/api/excel/excel.range#getspillparentornullobject--)|Gets the range object containing the anchor cell for a cell getting spilled into. Read-only.|
||[getSpillingToRange()](/javascript/api/excel/excel.range#getspillingtorange--)|Gets the range object containing the spill range when called on an anchor cell. Fails if applied to a range with more than one cell. Read-only.|
||[getSpillingToRangeOrNullObject()](/javascript/api/excel/excel.range#getspillingtorangeornullobject--)|Gets the range object containing the spill range when called on an anchor cell. Read-only.|
||[group(groupOption: Excel.GroupOption)](/javascript/api/excel/excel.range#group-groupoption-)|Groups columns and rows for an outline.|
||[hideGroupDetails(groupOption: Excel.GroupOption)](/javascript/api/excel/excel.range#hidegroupdetails-groupoption-)|Hide details of the row or column group.|
||[hasSpill](/javascript/api/excel/excel.range#hasspill)|Represents if all cells have a spill border.|
||[height](/javascript/api/excel/excel.range#height)|Returns the distance in points, for 100% zoom, from top edge of the range to bottom edge of the range. Read-only.|
||[left](/javascript/api/excel/excel.range#left)|Returns the distance in points, for 100% zoom, from left edge of the worksheet to left edge of the range. Read-only.|
||[savedAsArray](/javascript/api/excel/excel.range#savedasarray)|Represents if ALL the cells would be saved as an array formula.|
||[top](/javascript/api/excel/excel.range#top)|Returns the distance in points, for 100% zoom, from top edge of the worksheet to top edge of the range. Read-only.|
||[width](/javascript/api/excel/excel.range#width)|Returns the distance in points, for 100% zoom, from left edge of the range to right edge of the range. Read-only.|
||[showGroupDetails(groupOption: Excel.GroupOption)](/javascript/api/excel/excel.range#showgroupdetails-groupoption-)|Show details of the row or column group.|
||[ungroup(groupOption: Excel.GroupOption)](/javascript/api/excel/excel.range#ungroup-groupoption-)|Ungroups columns and rows for an outline.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[adjustIndent(amount: number)](/javascript/api/excel/excel.rangeformat#adjustindent-amount-)|Adjusts the indentation of the range formatting. The indent value ranges from 0 to 250.|
|[Shape](/javascript/api/excel/excel.shape)|[copyTo(destinationSheet?: Worksheet \| string)](/javascript/api/excel/excel.shape#copyto-destinationsheet-)|Copies and pastes a Shape object.|
||[placement](/javascript/api/excel/excel.shape#placement)|Represents how the object is attached to the cells below it.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|Creates a scalable vector graphic (SVG) from an XML string and adds it to the worksheet. Returns a Shape object that represents the new image.|
|[Slicer](/javascript/api/excel/excel.slicer)|[caption](/javascript/api/excel/excel.slicer#caption)|Represents the caption of slicer.|
||[clearFilters()](/javascript/api/excel/excel.slicer#clearfilters--)|Clears all the filters currently applied on the slicer.|
||[delete()](/javascript/api/excel/excel.slicer#delete--)|Deletes the slicer.|
||[getSelectedItems()](/javascript/api/excel/excel.slicer#getselecteditems--)|Returns an array of selected items' keys. Read-only.|
||[height](/javascript/api/excel/excel.slicer#height)|Represents the height, in points, of the slicer.|
||[left](/javascript/api/excel/excel.slicer#left)|Represents the distance, in points, from the left side of the slicer to the left of the worksheet.|
||[name](/javascript/api/excel/excel.slicer#name)|Represents the name of slicer.|
||[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|Represents the slicer name used in the formula.|
||[id](/javascript/api/excel/excel.slicer#id)|Represents the unique id of slicer. Read-only.|
||[isFilterCleared](/javascript/api/excel/excel.slicer#isfiltercleared)|True if all filters currently applied on the slicer are cleared.|
||[slicerItems](/javascript/api/excel/excel.slicer#sliceritems)|Represents the collection of SlicerItems that are part of the slicer. Read-only.|
||[worksheet](/javascript/api/excel/excel.slicer#worksheet)|Represents the worksheet containing the slicer. Read-only.|
||[selectItems(items?: string[])](/javascript/api/excel/excel.slicer#selectitems-items-)|Selects slicer items based on their keys. The previous selections are cleared.|
||[sortBy](/javascript/api/excel/excel.slicer#sortby)|Represents the sort order of the items in the slicer. Possible values are: "DataSourceOrder", "Ascending", "Descending".|
||[style](/javascript/api/excel/excel.slicer#style)|Constant value that represents the Slicer style. Possible values are: "SlicerStyleLight1" through "SlicerStyleLight6", "TableStyleOther1" through "TableStyleOther2", "SlicerStyleDark1" through "SlicerStyleDark6". A custom user-defined style present in the workbook can also be specified.|
||[top](/javascript/api/excel/excel.slicer#top)|Represents the distance, in points, from the top edge of the slicer to the top of the worksheet.|
||[width](/javascript/api/excel/excel.slicer#width)|Represents the width, in points, of the slicer.|
|[SlicerCollection](/javascript/api/excel/excel.slicercollection)|[add(slicerSource: string \| PivotTable \| Table, sourceField: string \| PivotField \| number \| TableColumn, slicerDestination?: string \| Worksheet)](/javascript/api/excel/excel.slicercollection#add-slicersource--sourcefield--slicerdestination-)|Adds a new slicer to the workbook.|
||[getCount()](/javascript/api/excel/excel.slicercollection#getcount--)|Returns the number of slicers in the collection.|
||[getItem(key: string)](/javascript/api/excel/excel.slicercollection#getitem-key-)|Gets a slicer object using its name or id.|
||[getItemAt(index: number)](/javascript/api/excel/excel.slicercollection#getitemat-index-)|Gets a slicer based on its position in the collection.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.slicercollection#getitemornullobject-key-)|Gets a slicer using its name or id. If the slicer does not exist, will return a null object.|
||[items](/javascript/api/excel/excel.slicercollection#items)|Gets the loaded child items in this collection.|
|[SlicerItem](/javascript/api/excel/excel.sliceritem)|[isSelected](/javascript/api/excel/excel.sliceritem#isselected)|True if the slicer item is selected.|
||[hasData](/javascript/api/excel/excel.sliceritem#hasdata)|True if the slicer item has data.|
||[key](/javascript/api/excel/excel.sliceritem#key)|Represents the unique value representing the slicer item.|
||[name](/javascript/api/excel/excel.sliceritem#name)|Represents the title displayed in the UI.|
|[SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection)|[getCount()](/javascript/api/excel/excel.sliceritemcollection#getcount--)|Returns the number of slicer items in the slicer.|
||[getItem(key: string)](/javascript/api/excel/excel.sliceritemcollection#getitem-key-)|Gets a slicer item object using its key or name.|
||[getItemAt(index: number)](/javascript/api/excel/excel.sliceritemcollection#getitemat-index-)|Gets a slicer item based on its position in the collection.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.sliceritemcollection#getitemornullobject-key-)|Gets a slicer item using its key or name. If the slicer item does not exist, will return a null object.|
||[items](/javascript/api/excel/excel.sliceritemcollection#items)|Gets the loaded child items in this collection.|
|[SlicerStyle](/javascript/api/excel/excel.slicerstyle)|[delete()](/javascript/api/excel/excel.slicerstyle#delete--)|Deletes the SlicerStyle.|
||[duplicate()](/javascript/api/excel/excel.slicerstyle#duplicate--)|Creates a duplicate of this SlicerStyle with copies of all the style elements.|
||[name](/javascript/api/excel/excel.slicerstyle#name)|Gets the name of the SlicerStyle.|
||[readOnly](/javascript/api/excel/excel.slicerstyle#readonly)|Specifies whether this SlicerStyle object is read-only. Read-only.|
|[SlicerStyleCollection](/javascript/api/excel/excel.slicerstylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.slicerstylecollection#add-name--makeuniquename-)|Creates a blank SlicerStyle with the specified name.|
||[getCount()](/javascript/api/excel/excel.slicerstylecollection#getcount--)|Gets the number of slicer styles in the collection.|
||[getDefault()](/javascript/api/excel/excel.slicerstylecollection#getdefault--)|Gets the default SlicerStyle for the parent object's scope.|
||[getItem(name: string)](/javascript/api/excel/excel.slicerstylecollection#getitem-name-)|Gets a SlicerStyle by name.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.slicerstylecollection#getitemornullobject-name-)|Gets a SlicerStyle by name. If the SlicerStyle does not exist, will return a null object.|
||[items](/javascript/api/excel/excel.slicerstylecollection#items)|Gets the loaded child items in this collection.|
||[setDefault(newDefaultStyle: SlicerStyle \| string)](/javascript/api/excel/excel.slicerstylecollection#setdefault-newdefaultstyle-)|Sets the default SlicerStyle for use in the parent object's scope.|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|Changes the table to use the default table style.|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|Occurs when filter is applied on a specific table.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|Occurs when filter is applied on any table in a workbook, or a worksheet.|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|Represents the id of the table in which the filter is applied.|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|Represents the type of the event. See Excel.EventType for details.|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|Represents the id of the worksheet which contains the table.|
|[TableStyle](/javascript/api/excel/excel.tablestyle)|[delete()](/javascript/api/excel/excel.tablestyle#delete--)|Deletes the TableStyle.|
||[duplicate()](/javascript/api/excel/excel.tablestyle#duplicate--)|Creates a duplicate of this TableStyle with copies of all the style elements.|
||[name](/javascript/api/excel/excel.tablestyle#name)|Gets the name of the TableStyle.|
||[readOnly](/javascript/api/excel/excel.tablestyle#readonly)|Specifies whether this TableStyle object is read-only. Read-only.|
|[TableStyleCollection](/javascript/api/excel/excel.tablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.tablestylecollection#add-name--makeuniquename-)|Creates a blank TableStyle with the specified name.|
||[getCount()](/javascript/api/excel/excel.tablestylecollection#getcount--)|Gets the number of table styles in the collection.|
||[getDefault()](/javascript/api/excel/excel.tablestylecollection#getdefault--)|Gets the default TableStyle for the parent object's scope.|
||[getItem(name: string)](/javascript/api/excel/excel.tablestylecollection#getitem-name-)|Gets a TableStyle by name.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.tablestylecollection#getitemornullobject-name-)|Gets a TableStyle by name. If the TableStyle does not exist, will return a null object.|
||[items](/javascript/api/excel/excel.tablestylecollection#items)|Gets the loaded child items in this collection.|
||[setDefault(newDefaultStyle: TableStyle \| string)](/javascript/api/excel/excel.tablestylecollection#setdefault-newdefaultstyle-)|Sets the default TableStyle for use in the parent object's scope.|
|[TimelineStyle](/javascript/api/excel/excel.timelinestyle)|[delete()](/javascript/api/excel/excel.timelinestyle#delete--)|Deletes the TableStyle.|
||[duplicate()](/javascript/api/excel/excel.timelinestyle#duplicate--)|Creates a duplicate of this TimelineStyle with copies of all the style elements.|
||[name](/javascript/api/excel/excel.timelinestyle#name)|Gets the name of the TimelineStyle.|
||[readOnly](/javascript/api/excel/excel.timelinestyle#readonly)|Specifies whether this TimelineStyle object is read-only. Read-only.|
|[TimelineStyleCollection](/javascript/api/excel/excel.timelinestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.timelinestylecollection#add-name--makeuniquename-)|Creates a blank TimelineStyle with the specified name.|
||[getCount()](/javascript/api/excel/excel.timelinestylecollection#getcount--)|Gets the number of timeline styles in the collection.|
||[getDefault()](/javascript/api/excel/excel.timelinestylecollection#getdefault--)|Gets the default TimelineStyle for the parent object's scope.|
||[getItem(name: string)](/javascript/api/excel/excel.timelinestylecollection#getitem-name-)|Gets a TimelineStyle by name.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.timelinestylecollection#getitemornullobject-name-)|Gets a TimelineStyle by name. If the TimelineStyle does not exist, will return a null object.|
||[items](/javascript/api/excel/excel.timelinestylecollection#items)|Gets the loaded child items in this collection.|
||[setDefault(newDefaultStyle: TimelineStyle \| string)](/javascript/api/excel/excel.timelinestylecollection#setdefault-newdefaultstyle-)|Sets the default TimelineStyle for use in the parent object's scope.|
|[Workbook](/javascript/api/excel/excel.workbook)|[close(closeBehavior?: Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#close-closebehavior-)|Close current workbook.|
||[getActiveSlicer()](/javascript/api/excel/excel.workbook#getactiveslicer--)|Gets the currently active slicer in the workbook. If there is no active slicer, an `ItemNotFound` exception is thrown.|
||[getActiveSlicerOrNullObject()](/javascript/api/excel/excel.workbook#getactiveslicerornullobject--)|Gets the currently active slicer in the workbook. If there is no active slicer, a null object is returned.|
||[comments](/javascript/api/excel/excel.workbook#comments)|Represents a collection of Comments associated with the workbook. Read-only.|
||[pivotTableStyles](/javascript/api/excel/excel.workbook#pivottablestyles)|Represents a collection of PivotTableStyles associated with the workbook. Read-only.|
||[slicerStyles](/javascript/api/excel/excel.workbook#slicerstyles)|Represents a collection of SlicerStyles associated with the workbook. Read-only.|
||[slicers](/javascript/api/excel/excel.workbook#slicers)|Represents a collection of Slicers associated with the workbook. Read-only.|
||[tableStyles](/javascript/api/excel/excel.workbook#tablestyles)|Represents a collection of TableStyles associated with the workbook. Read-only.|
||[timelineStyles](/javascript/api/excel/excel.workbook#timelinestyles)|Represents a collection of TimelineStyles associated with the workbook. Read-only.|
||[save(saveBehavior?: Excel.SaveBehavior)](/javascript/api/excel/excel.workbook#save-savebehavior-)|Save current workbook.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|True if the workbook uses the 1904 date system.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[comments](/javascript/api/excel/excel.worksheet#comments)|Returns a collection of all the Comments objects on the worksheet. Read-only.|
||[onColumnSorted](/javascript/api/excel/excel.worksheet#oncolumnsorted)|Occurs when one or more columns have been sorted. This happens as the result of a left-to-right sort operation.|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|Occurs when filter is applied on a specific worksheet.|
||[onRowHiddenChanged](/javascript/api/excel/excel.worksheet#onrowhiddenchanged)|Occurs when the hidden state of one or more rows has changed on a specific worksheet.|
||[onRowSorted](/javascript/api/excel/excel.worksheet#onrowsorted)|Occurs when one or more rows have been sorted. This happens as the result of a top-to-bottom sort operation.|
||[onSingleClicked](/javascript/api/excel/excel.worksheet#onsingleclicked)|Occurs when left-clicked/tapped operation happens in the worksheet. This event will not be fired when clicking in the following cases:|
||[slicers](/javascript/api/excel/excel.worksheet#slicers)|Returns collection of slicers that are part of the worksheet. Read-only.|
||[showOutlineLevels(rowLevels: number, columnLevels: number)](/javascript/api/excel/excel.worksheet#showoutlinelevels-rowlevels--columnlevels-)|Shows row or column groups by their outline levels.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|Inserts the specified worksheets of a workbook into the current workbook.|
||[onColumnSorted](/javascript/api/excel/excel.worksheetcollection#oncolumnsorted)|Occurs when one or more columns have been sorted. This happens as the result of a left-to-right sort operation.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|Occurs when any worksheet's filter is applied in the workbook.|
||[onRowHiddenChanged](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged)|Occurs when the hidden state of one or more rows has changed on a specific worksheet.|
||[onRowSorted](/javascript/api/excel/excel.worksheetcollection#onrowsorted)|Occurs when one or more rows have been sorted. This happens as the result of a top-to-bottom sort operation.|
||[onSingleClicked](/javascript/api/excel/excel.worksheetcollection#onsingleclicked)|Occurs when left-clicked/tapped operation happens in the worksheet collection. This event will not be fired when clicking in the following cases:|
|[WorksheetColumnSortedEventArgs](/javascript/api/excel/excel.worksheetcolumnsortedeventargs)|[address](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#address)|Gets the range address that represents the sorted areas of a specific worksheet. Only columns changed as a result of the sort operation are returned.|
||[source](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#source)|Gets the source of the event. See Excel.EventSource for details.|
||[type](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#type)|Gets the type of the event. See Excel.EventType for details.|
||[worksheetId](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#worksheetid)|Gets the id of the worksheet where the sorting happened.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|Represents the type of the event. See Excel.EventType for details.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|Represents the id of the worksheet in which the filter is applied.|
|[WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs)|[address](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#address)|Gets the range address that represents the changed area of a specific worksheet.|
||[changeType](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#changetype)|Gets the type of change that represents how the event was triggered. See `Excel.RowHiddenChangeType` for details.|
||[source](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#source)|Gets the source of the event. See Excel.EventSource for details.|
||[type](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#type)|Gets the type of the event. See Excel.EventType for details.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#worksheetid)|Gets the id of the worksheet in which the data changed.|
|[WorksheetRowSortedEventArgs](/javascript/api/excel/excel.worksheetrowsortedeventargs)|[address](/javascript/api/excel/excel.worksheetrowsortedeventargs#address)|Gets the range address that represents the sorted areas of a specific worksheet. Only rows changed as a result of the sort operation are returned.|
||[source](/javascript/api/excel/excel.worksheetrowsortedeventargs#source)|Gets the source of the event. See Excel.EventSource for details.|
||[type](/javascript/api/excel/excel.worksheetrowsortedeventargs#type)|Gets the type of the event. See Excel.EventType for details.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowsortedeventargs#worksheetid)|Gets the id of the worksheet where the sorting happened.|
|[WorksheetSingleClickedEventArgs](/javascript/api/excel/excel.worksheetsingleclickedeventargs)|[address](/javascript/api/excel/excel.worksheetsingleclickedeventargs#address)|Gets the address that represents the cell which was left-clicked/tapped for a specific worksheet.|
||[offsetX](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsetx)|The distance, in points, from the left-clicked/tapped point to the left (or right for right-to-left languages) gridline edge of the left-clicked/tapped cell.|
||[offsetY](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsety)|The distance, in points, from the left-clicked/tapped point to the top gridline edge of the left-clicked/tapped cell.|
||[type](/javascript/api/excel/excel.worksheetsingleclickedeventargs#type)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetsingleclickedeventargs#worksheetid)|Gets the id of the worksheet in which the cell was left-clicked/tapped.|

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel?view=excel-js-preview)
- [Excel JavaScript API requirement sets](./excel-api-requirement-sets.md)
