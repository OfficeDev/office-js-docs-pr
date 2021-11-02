---
title: Excel JavaScript API requirement set 1.10
description: 'Details about the ExcelApi 1.10 requirement set.'
ms.date: 04/02/2021
ms.prod: excel
ms.localizationpriority: medium
---

# What's new in Excel JavaScript API 1.10

The ExcelApi 1.10 introduced key features, such as commenting, outlines, and slicers. It also added event support for worksheet-level clicking and sorting.

| Feature area | Description | Relevant objects |
|:--- |:--- |:--- |
| [Comments](../../excel/excel-add-ins-comments.md) | Add, edit, and delete comments. | [Comment](/javascript/api/excel/excel.comment), [CommentCollection](/javascript/api/excel/excel.commentcollection) |
| [Outlines](../../excel/excel-add-ins-ranges-group.md) | Group rows and columns to form collapsible outlines. | [Range](/javascript/api/excel/excel.range), [Worksheet](/javascript/api/excel/excel.worksheet) |
| [Slicers](../../excel/excel-add-ins-pivottables.md#filter-with-slicers) | Insert and configure slicers to tables and PivotTables. | [Slicer](/javascript/api/excel/excel.slicer) |
| [More Worksheet Events](../../excel/excel-add-ins-events.md) | Listen for click and sort events in the worksheet. | [Worksheet (Events)](/javascript/api/excel/excel.worksheet#events) |

## API list

The following table lists the APIs in Excel JavaScript API requirement set 1.10. To view API reference documentation for all APIs supported by Excel JavaScript API requirement set 1.10 or earlier, see [Excel APIs in requirement set 1.10 or earlier](/javascript/api/excel?view=excel-js-1.10&preserve-view=true).

| Class | Fields | Description |
|:---|:---|:---|
|[Comment](/javascript/api/excel/excel.comment)|[authorEmail](/javascript/api/excel/excel.comment#authorEmail)|Gets the email of the comment's author.|
||[authorName](/javascript/api/excel/excel.comment#authorName)|Gets the name of the comment's author.|
||[content](/javascript/api/excel/excel.comment#content)|The comment's content.|
||[creationDate](/javascript/api/excel/excel.comment#creationDate)|Gets the creation time of the comment.|
||[delete()](/javascript/api/excel/excel.comment#delete__)|Deletes the comment and all the connected replies.|
||[getLocation()](/javascript/api/excel/excel.comment#getLocation__)|Gets the cell where this comment is located.|
||[id](/javascript/api/excel/excel.comment#id)|Specifies the comment identifier.|
||[replies](/javascript/api/excel/excel.comment#replies)|Represents a collection of reply objects associated with the comment.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add(cellAddress: Range \| string, content: string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentcollection#add_cellAddress__content__contentType_)|Creates a new comment with the given content on the given cell.|
||[getCount()](/javascript/api/excel/excel.commentcollection#getCount__)|Gets the number of comments in the collection.|
||[getItem(commentId: string)](/javascript/api/excel/excel.commentcollection#getItem_commentId_)|Gets a comment from the collection based on its ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentcollection#getItemAt_index_)|Gets a comment from the collection based on its position.|
||[getItemByCell(cellAddress: Range \| string)](/javascript/api/excel/excel.commentcollection#getItemByCell_cellAddress_)|Gets the comment from the specified cell.|
||[getItemByReplyId(replyId: string)](/javascript/api/excel/excel.commentcollection#getItemByReplyId_replyId_)|Gets the comment to which the given reply is connected.|
||[items](/javascript/api/excel/excel.commentcollection#items)|Gets the loaded child items in this collection.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[authorEmail](/javascript/api/excel/excel.commentreply#authorEmail)|Gets the email of the comment reply's author.|
||[authorName](/javascript/api/excel/excel.commentreply#authorName)|Gets the name of the comment reply's author.|
||[content](/javascript/api/excel/excel.commentreply#content)|The comment reply's content.|
||[creationDate](/javascript/api/excel/excel.commentreply#creationDate)|Gets the creation time of the comment reply.|
||[delete()](/javascript/api/excel/excel.commentreply#delete__)|Deletes the comment reply.|
||[getLocation()](/javascript/api/excel/excel.commentreply#getLocation__)|Gets the cell where this comment reply is located.|
||[getParentComment()](/javascript/api/excel/excel.commentreply#getParentComment__)|Gets the parent comment of this reply.|
||[id](/javascript/api/excel/excel.commentreply#id)|Specifies the comment reply identifier.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add(content: string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentreplycollection#add_content__contentType_)|Creates a comment reply for a comment.|
||[getCount()](/javascript/api/excel/excel.commentreplycollection#getCount__)|Gets the number of comment replies in the collection.|
||[getItem(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#getItem_commentReplyId_)|Returns a comment reply identified by its ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentreplycollection#getItemAt_index_)|Gets a comment reply based on its position in the collection.|
||[items](/javascript/api/excel/excel.commentreplycollection#items)|Gets the loaded child items in this collection.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[enableFieldList](/javascript/api/excel/excel.pivotlayout#enableFieldList)|Specifies if the field list can be shown in the UI.|
|[PivotTableStyle](/javascript/api/excel/excel.pivottablestyle)|[delete()](/javascript/api/excel/excel.pivottablestyle#delete__)|Deletes the PivotTable style.|
||[duplicate()](/javascript/api/excel/excel.pivottablestyle#duplicate__)|Creates a duplicate of this PivotTable style with copies of all the style elements.|
||[name](/javascript/api/excel/excel.pivottablestyle#name)|Gets the name of the PivotTable style.|
||[readOnly](/javascript/api/excel/excel.pivottablestyle#readOnly)|Specifies if this `PivotTableStyle` object is read-only.|
|[PivotTableStyleCollection](/javascript/api/excel/excel.pivottablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.pivottablestylecollection#add_name__makeUniqueName_)|Creates a blank `PivotTableStyle` with the specified name.|
||[getCount()](/javascript/api/excel/excel.pivottablestylecollection#getCount__)|Gets the number of PivotTable styles in the collection.|
||[getDefault()](/javascript/api/excel/excel.pivottablestylecollection#getDefault__)|Gets the default PivotTable style for the parent object's scope.|
||[getItem(name: string)](/javascript/api/excel/excel.pivottablestylecollection#getItem_name_)|Gets a `PivotTableStyle` by name.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablestylecollection#getItemOrNullObject_name_)|Gets a `PivotTableStyle` by name.|
||[items](/javascript/api/excel/excel.pivottablestylecollection#items)|Gets the loaded child items in this collection.|
||[setDefault(newDefaultStyle: PivotTableStyle \| string)](/javascript/api/excel/excel.pivottablestylecollection#setDefault_newDefaultStyle_)|Sets the default PivotTable style for use in the parent object's scope.|
|[Range](/javascript/api/excel/excel.range)|[group(groupOption: Excel.GroupOption)](/javascript/api/excel/excel.range#group_groupOption_)|Groups columns and rows for an outline.|
||[height](/javascript/api/excel/excel.range#height)|Returns the distance in points, for 100% zoom, from the top edge of the range to the bottom edge of the range.|
||[hideGroupDetails(groupOption: Excel.GroupOption)](/javascript/api/excel/excel.range#hideGroupDetails_groupOption_)|Hides the details of the row or column group.|
||[left](/javascript/api/excel/excel.range#left)|Returns the distance in points, for 100% zoom, from the left edge of the worksheet to the left edge of the range.|
||[showGroupDetails(groupOption: Excel.GroupOption)](/javascript/api/excel/excel.range#showGroupDetails_groupOption_)|Shows the details of the row or column group.|
||[top](/javascript/api/excel/excel.range#top)|Returns the distance in points, for 100% zoom, from the top edge of the worksheet to the top edge of the range.|
||[ungroup(groupOption: Excel.GroupOption)](/javascript/api/excel/excel.range#ungroup_groupOption_)|Ungroups columns and rows for an outline.|
||[width](/javascript/api/excel/excel.range#width)|Returns the distance in points, for 100% zoom, from the left edge of the range to the right edge of the range.|
|[Shape](/javascript/api/excel/excel.shape)|[copyTo(destinationSheet?: Worksheet \| string)](/javascript/api/excel/excel.shape#copyTo_destinationSheet_)|Copies and pastes a `Shape` object.|
||[placement](/javascript/api/excel/excel.shape#placement)|Represents how the object is attached to the cells below it.|
|[Slicer](/javascript/api/excel/excel.slicer)|[caption](/javascript/api/excel/excel.slicer#caption)|Represents the caption of the slicer.|
||[clearFilters()](/javascript/api/excel/excel.slicer#clearFilters__)|Clears all the filters currently applied on the slicer.|
||[delete()](/javascript/api/excel/excel.slicer#delete__)|Deletes the slicer.|
||[getSelectedItems()](/javascript/api/excel/excel.slicer#getSelectedItems__)|Returns an array of selected items' keys.|
||[height](/javascript/api/excel/excel.slicer#height)|Represents the height, in points, of the slicer.|
||[id](/javascript/api/excel/excel.slicer#id)|Represents the unique ID of the slicer.|
||[isFilterCleared](/javascript/api/excel/excel.slicer#isFilterCleared)|Value is `true` if all filters currently applied on the slicer are cleared.|
||[left](/javascript/api/excel/excel.slicer#left)|Represents the distance, in points, from the left side of the slicer to the left of the worksheet.|
||[name](/javascript/api/excel/excel.slicer#name)|Represents the name of the slicer.|
||[selectItems(items?: string[])](/javascript/api/excel/excel.slicer#selectItems_items_)|Selects slicer items based on their keys.|
||[slicerItems](/javascript/api/excel/excel.slicer#slicerItems)|Represents the collection of slicer items that are part of the slicer.|
||[sortBy](/javascript/api/excel/excel.slicer#sortBy)|Represents the sort order of the items in the slicer.|
||[style](/javascript/api/excel/excel.slicer#style)|Constant value that represents the slicer style.|
||[top](/javascript/api/excel/excel.slicer#top)|Represents the distance, in points, from the top edge of the slicer to the top of the worksheet.|
||[width](/javascript/api/excel/excel.slicer#width)|Represents the width, in points, of the slicer.|
||[worksheet](/javascript/api/excel/excel.slicer#worksheet)|Represents the worksheet containing the slicer.|
|[SlicerCollection](/javascript/api/excel/excel.slicercollection)|[add(slicerSource: string \| PivotTable \| Table, sourceField: string \| PivotField \| number \| TableColumn, slicerDestination?: string \| Worksheet)](/javascript/api/excel/excel.slicercollection#add_slicerSource__sourceField__slicerDestination_)|Adds a new slicer to the workbook.|
||[getCount()](/javascript/api/excel/excel.slicercollection#getCount__)|Returns the number of slicers in the collection.|
||[getItem(key: string)](/javascript/api/excel/excel.slicercollection#getItem_key_)|Gets a slicer object using its name or ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.slicercollection#getItemAt_index_)|Gets a slicer based on its position in the collection.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.slicercollection#getItemOrNullObject_key_)|Gets a slicer using its name or ID.|
||[items](/javascript/api/excel/excel.slicercollection#items)|Gets the loaded child items in this collection.|
|[SlicerItem](/javascript/api/excel/excel.sliceritem)|[hasData](/javascript/api/excel/excel.sliceritem#hasData)|Value is `true` if the slicer item has data.|
||[isSelected](/javascript/api/excel/excel.sliceritem#isSelected)|Value is `true` if the slicer item is selected.|
||[key](/javascript/api/excel/excel.sliceritem#key)|Represents the unique value representing the slicer item.|
||[name](/javascript/api/excel/excel.sliceritem#name)|Represents the title displayed in the Excel UI.|
|[SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection)|[getCount()](/javascript/api/excel/excel.sliceritemcollection#getCount__)|Returns the number of slicer items in the slicer.|
||[getItem(key: string)](/javascript/api/excel/excel.sliceritemcollection#getItem_key_)|Gets a slicer item object using its key or name.|
||[getItemAt(index: number)](/javascript/api/excel/excel.sliceritemcollection#getItemAt_index_)|Gets a slicer item based on its position in the collection.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.sliceritemcollection#getItemOrNullObject_key_)|Gets a slicer item using its key or name.|
||[items](/javascript/api/excel/excel.sliceritemcollection#items)|Gets the loaded child items in this collection.|
|[SlicerStyle](/javascript/api/excel/excel.slicerstyle)|[delete()](/javascript/api/excel/excel.slicerstyle#delete__)|Deletes the slicer style.|
||[duplicate()](/javascript/api/excel/excel.slicerstyle#duplicate__)|Creates a duplicate of this slicer style with copies of all the style elements.|
||[name](/javascript/api/excel/excel.slicerstyle#name)|Gets the name of the slicer style.|
||[readOnly](/javascript/api/excel/excel.slicerstyle#readOnly)|Specifies if this `SlicerStyle` object is read-only.|
|[SlicerStyleCollection](/javascript/api/excel/excel.slicerstylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.slicerstylecollection#add_name__makeUniqueName_)|Creates a blank slicer style with the specified name.|
||[getCount()](/javascript/api/excel/excel.slicerstylecollection#getCount__)|Gets the number of slicer styles in the collection.|
||[getDefault()](/javascript/api/excel/excel.slicerstylecollection#getDefault__)|Gets the default `SlicerStyle` for the parent object's scope.|
||[getItem(name: string)](/javascript/api/excel/excel.slicerstylecollection#getItem_name_)|Gets a `SlicerStyle` by name.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.slicerstylecollection#getItemOrNullObject_name_)|Gets a `SlicerStyle` by name.|
||[items](/javascript/api/excel/excel.slicerstylecollection#items)|Gets the loaded child items in this collection.|
||[setDefault(newDefaultStyle: SlicerStyle \| string)](/javascript/api/excel/excel.slicerstylecollection#setDefault_newDefaultStyle_)|Sets the default slicer style for use in the parent object's scope.|
|[TableStyle](/javascript/api/excel/excel.tablestyle)|[delete()](/javascript/api/excel/excel.tablestyle#delete__)|Deletes the table style.|
||[duplicate()](/javascript/api/excel/excel.tablestyle#duplicate__)|Creates a duplicate of this table style with copies of all the style elements.|
||[name](/javascript/api/excel/excel.tablestyle#name)|Gets the name of the table style.|
||[readOnly](/javascript/api/excel/excel.tablestyle#readOnly)|Specifies if this `TableStyle` object is read-only.|
|[TableStyleCollection](/javascript/api/excel/excel.tablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.tablestylecollection#add_name__makeUniqueName_)|Creates a blank `TableStyle` with the specified name.|
||[getCount()](/javascript/api/excel/excel.tablestylecollection#getCount__)|Gets the number of table styles in the collection.|
||[getDefault()](/javascript/api/excel/excel.tablestylecollection#getDefault__)|Gets the default table style for the parent object's scope.|
||[getItem(name: string)](/javascript/api/excel/excel.tablestylecollection#getItem_name_)|Gets a `TableStyle` by name.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.tablestylecollection#getItemOrNullObject_name_)|Gets a `TableStyle` by name.|
||[items](/javascript/api/excel/excel.tablestylecollection#items)|Gets the loaded child items in this collection.|
||[setDefault(newDefaultStyle: TableStyle \| string)](/javascript/api/excel/excel.tablestylecollection#setDefault_newDefaultStyle_)|Sets the default table style for use in the parent object's scope.|
|[TimelineStyle](/javascript/api/excel/excel.timelinestyle)|[delete()](/javascript/api/excel/excel.timelinestyle#delete__)|Deletes the table style.|
||[duplicate()](/javascript/api/excel/excel.timelinestyle#duplicate__)|Creates a duplicate of this timeline style with copies of all the style elements.|
||[name](/javascript/api/excel/excel.timelinestyle#name)|Gets the name of the timeline style.|
||[readOnly](/javascript/api/excel/excel.timelinestyle#readOnly)|Specifies if this `TimelineStyle` object is read-only.|
|[TimelineStyleCollection](/javascript/api/excel/excel.timelinestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.timelinestylecollection#add_name__makeUniqueName_)|Creates a blank `TimelineStyle` with the specified name.|
||[getCount()](/javascript/api/excel/excel.timelinestylecollection#getCount__)|Gets the number of timeline styles in the collection.|
||[getDefault()](/javascript/api/excel/excel.timelinestylecollection#getDefault__)|Gets the default timeline style for the parent object's scope.|
||[getItem(name: string)](/javascript/api/excel/excel.timelinestylecollection#getItem_name_)|Gets a `TimelineStyle` by name.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.timelinestylecollection#getItemOrNullObject_name_)|Gets a `TimelineStyle` by name.|
||[items](/javascript/api/excel/excel.timelinestylecollection#items)|Gets the loaded child items in this collection.|
||[setDefault(newDefaultStyle: TimelineStyle \| string)](/javascript/api/excel/excel.timelinestylecollection#setDefault_newDefaultStyle_)|Sets the default timeline style for use in the parent object's scope.|
|[Workbook](/javascript/api/excel/excel.workbook)|[comments](/javascript/api/excel/excel.workbook#comments)|Represents a collection of comments associated with the workbook.|
||[getActiveSlicer()](/javascript/api/excel/excel.workbook#getActiveSlicer__)|Gets the currently active slicer in the workbook.|
||[getActiveSlicerOrNullObject()](/javascript/api/excel/excel.workbook#getActiveSlicerOrNullObject__)|Gets the currently active slicer in the workbook.|
||[pivotTableStyles](/javascript/api/excel/excel.workbook#pivotTableStyles)|Represents a collection of PivotTableStyles associated with the workbook.|
||[slicerStyles](/javascript/api/excel/excel.workbook#slicerStyles)|Represents a collection of SlicerStyles associated with the workbook.|
||[slicers](/javascript/api/excel/excel.workbook#slicers)|Represents a collection of slicers associated with the workbook.|
||[tableStyles](/javascript/api/excel/excel.workbook#tableStyles)|Represents a collection of TableStyles associated with the workbook.|
||[timelineStyles](/javascript/api/excel/excel.workbook#timelineStyles)|Represents a collection of TimelineStyles associated with the workbook.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[comments](/javascript/api/excel/excel.worksheet#comments)|Returns a collection of all the Comments objects on the worksheet.|
||[onColumnSorted](/javascript/api/excel/excel.worksheet#onColumnSorted)|Occurs when one or more columns have been sorted.|
||[onRowSorted](/javascript/api/excel/excel.worksheet#onRowSorted)|Occurs when one or more rows have been sorted.|
||[onSingleClicked](/javascript/api/excel/excel.worksheet#onSingleClicked)|Occurs when a left-clicked/tapped action happens in the worksheet.|
||[showOutlineLevels(rowLevels: number, columnLevels: number)](/javascript/api/excel/excel.worksheet#showOutlineLevels_rowLevels__columnLevels_)|Shows row or column groups by their outline levels.|
||[slicers](/javascript/api/excel/excel.worksheet#slicers)|Returns a collection of slicers that are part of the worksheet.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onColumnSorted](/javascript/api/excel/excel.worksheetcollection#onColumnSorted)|Occurs when one or more columns have been sorted.|
||[onRowSorted](/javascript/api/excel/excel.worksheetcollection#onRowSorted)|Occurs when one or more rows have been sorted.|
||[onSingleClicked](/javascript/api/excel/excel.worksheetcollection#onSingleClicked)|Occurs when left-clicked/tapped operation happens in the worksheet collection.|
|[WorksheetColumnSortedEventArgs](/javascript/api/excel/excel.worksheetcolumnsortedeventargs)|[address](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#address)|Gets the range address that represents the sorted areas of a specific worksheet.|
||[source](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#source)|Gets the source of the event.|
||[type](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#type)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#worksheetId)|Gets the ID of the worksheet where the sorting happened.|
|[WorksheetRowSortedEventArgs](/javascript/api/excel/excel.worksheetrowsortedeventargs)|[address](/javascript/api/excel/excel.worksheetrowsortedeventargs#address)|Gets the range address that represents the sorted areas of a specific worksheet.|
||[source](/javascript/api/excel/excel.worksheetrowsortedeventargs#source)|Gets the source of the event.|
||[type](/javascript/api/excel/excel.worksheetrowsortedeventargs#type)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowsortedeventargs#worksheetId)|Gets the ID of the worksheet where the sorting happened.|
|[WorksheetSingleClickedEventArgs](/javascript/api/excel/excel.worksheetsingleclickedeventargs)|[address](/javascript/api/excel/excel.worksheetsingleclickedeventargs#address)|Gets the address that represents the cell which was left-clicked/tapped for a specific worksheet.|
||[offsetX](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsetX)|The distance, in points, from the left-clicked/tapped point to the left (or right for right-to-left languages) gridline edge of the left-clicked/tapped cell.|
||[offsetY](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsetY)|The distance, in points, from the left-clicked/tapped point to the top gridline edge of the left-clicked/tapped cell.|
||[type](/javascript/api/excel/excel.worksheetsingleclickedeventargs#type)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetsingleclickedeventargs#worksheetId)|Gets the ID of the worksheet in which the cell was left-clicked/tapped.|

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel?view=excel-js-1.10&preserve-view=true)
- [Excel JavaScript API requirement sets](excel-api-requirement-sets.md)