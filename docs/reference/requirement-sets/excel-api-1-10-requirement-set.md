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
| [More Worksheet Events](../../excel/excel-add-ins-events.md) | Listen for click and sort events in the worksheet. | [Worksheet (Events)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-events-member) |

## API list

The following table lists the APIs in Excel JavaScript API requirement set 1.10. To view API reference documentation for all APIs supported by Excel JavaScript API requirement set 1.10 or earlier, see [Excel APIs in requirement set 1.10 or earlier](/javascript/api/excel?view=excel-js-1.10&preserve-view=true).

| Class | Fields | Description |
|:---|:---|:---|
|[Comment](/javascript/api/excel/excel.comment)|[authorEmail](/javascript/api/excel/excel.comment#excel-excel-comment-authoremail-member)|Gets the email of the comment's author.|
||[authorName](/javascript/api/excel/excel.comment#excel-excel-comment-authorname-member)|Gets the name of the comment's author.|
||[content](/javascript/api/excel/excel.comment#excel-excel-comment-content-member)|The comment's content.|
||[creationDate](/javascript/api/excel/excel.comment#excel-excel-comment-creationdate-member)|Gets the creation time of the comment.|
||[delete()](/javascript/api/excel/excel.comment#excel-excel-comment-delete-member(1))|Deletes the comment and all the connected replies.|
||[getLocation()](/javascript/api/excel/excel.comment#excel-excel-comment-getlocation-member(1))|Gets the cell where this comment is located.|
||[id](/javascript/api/excel/excel.comment#excel-excel-comment-id-member)|Specifies the comment identifier.|
||[replies](/javascript/api/excel/excel.comment#excel-excel-comment-replies-member)|Represents a collection of reply objects associated with the comment.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add(cellAddress: Range \| string, content: string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-add-member(1))|Creates a new comment with the given content on the given cell.|
||[getCount()](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-getcount-member(1))|Gets the number of comments in the collection.|
||[getItem(commentId: string)](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-getitem-member(1))|Gets a comment from the collection based on its ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-getitemat-member(1))|Gets a comment from the collection based on its position.|
||[getItemByCell(cellAddress: Range \| string)](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-getitembycell-member(1))|Gets the comment from the specified cell.|
||[getItemByReplyId(replyId: string)](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-getitembyreplyid-member(1))|Gets the comment to which the given reply is connected.|
||[items](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-items-member)|Gets the loaded child items in this collection.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[authorEmail](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-authoremail-member)|Gets the email of the comment reply's author.|
||[authorName](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-authorname-member)|Gets the name of the comment reply's author.|
||[content](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-content-member)|The comment reply's content.|
||[creationDate](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-creationdate-member)|Gets the creation time of the comment reply.|
||[delete()](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-delete-member(1))|Deletes the comment reply.|
||[getLocation()](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-getlocation-member(1))|Gets the cell where this comment reply is located.|
||[getParentComment()](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-getparentcomment-member(1))|Gets the parent comment of this reply.|
||[id](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-id-member)|Specifies the comment reply identifier.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add(content: string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-add-member(1))|Creates a comment reply for a comment.|
||[getCount()](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-getcount-member(1))|Gets the number of comment replies in the collection.|
||[getItem(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-getitem-member(1))|Returns a comment reply identified by its ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-getitemat-member(1))|Gets a comment reply based on its position in the collection.|
||[items](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-items-member)|Gets the loaded child items in this collection.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[enableFieldList](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-enablefieldlist-member)|Specifies if the field list can be shown in the UI.|
|[PivotTableStyle](/javascript/api/excel/excel.pivottablestyle)|[delete()](/javascript/api/excel/excel.pivottablestyle#excel-excel-pivottablestyle-delete-member(1))|Deletes the PivotTable style.|
||[duplicate()](/javascript/api/excel/excel.pivottablestyle#excel-excel-pivottablestyle-duplicate-member(1))|Creates a duplicate of this PivotTable style with copies of all the style elements.|
||[name](/javascript/api/excel/excel.pivottablestyle#excel-excel-pivottablestyle-name-member)|Gets the name of the PivotTable style.|
||[readOnly](/javascript/api/excel/excel.pivottablestyle#excel-excel-pivottablestyle-readonly-member)|Specifies if this `PivotTableStyle` object is read-only.|
|[PivotTableStyleCollection](/javascript/api/excel/excel.pivottablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.pivottablestylecollection#excel-excel-pivottablestylecollection-add-member(1))|Creates a blank `PivotTableStyle` with the specified name.|
||[getCount()](/javascript/api/excel/excel.pivottablestylecollection#excel-excel-pivottablestylecollection-getcount-member(1))|Gets the number of PivotTable styles in the collection.|
||[getDefault()](/javascript/api/excel/excel.pivottablestylecollection#excel-excel-pivottablestylecollection-getdefault-member(1))|Gets the default PivotTable style for the parent object's scope.|
||[getItem(name: string)](/javascript/api/excel/excel.pivottablestylecollection#excel-excel-pivottablestylecollection-getitem-member(1))|Gets a `PivotTableStyle` by name.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablestylecollection#excel-excel-pivottablestylecollection-getitemornullobject-member(1))|Gets a `PivotTableStyle` by name.|
||[items](/javascript/api/excel/excel.pivottablestylecollection#excel-excel-pivottablestylecollection-items-member)|Gets the loaded child items in this collection.|
||[setDefault(newDefaultStyle: PivotTableStyle \| string)](/javascript/api/excel/excel.pivottablestylecollection#excel-excel-pivottablestylecollection-setdefault-member(1))|Sets the default PivotTable style for use in the parent object's scope.|
|[Range](/javascript/api/excel/excel.range)|[group(groupOption: Excel.GroupOption)](/javascript/api/excel/excel.range#excel-excel-range-group-member(1))|Groups columns and rows for an outline.|
||[height](/javascript/api/excel/excel.range#excel-excel-range-height-member)|Returns the distance in points, for 100% zoom, from the top edge of the range to the bottom edge of the range.|
||[hideGroupDetails(groupOption: Excel.GroupOption)](/javascript/api/excel/excel.range#excel-excel-range-hidegroupdetails-member(1))|Hides the details of the row or column group.|
||[left](/javascript/api/excel/excel.range#excel-excel-range-left-member)|Returns the distance in points, for 100% zoom, from the left edge of the worksheet to the left edge of the range.|
||[showGroupDetails(groupOption: Excel.GroupOption)](/javascript/api/excel/excel.range#excel-excel-range-showgroupdetails-member(1))|Shows the details of the row or column group.|
||[top](/javascript/api/excel/excel.range#excel-excel-range-top-member)|Returns the distance in points, for 100% zoom, from the top edge of the worksheet to the top edge of the range.|
||[ungroup(groupOption: Excel.GroupOption)](/javascript/api/excel/excel.range#excel-excel-range-ungroup-member(1))|Ungroups columns and rows for an outline.|
||[width](/javascript/api/excel/excel.range#excel-excel-range-width-member)|Returns the distance in points, for 100% zoom, from the left edge of the range to the right edge of the range.|
|[Shape](/javascript/api/excel/excel.shape)|[copyTo(destinationSheet?: Worksheet \| string)](/javascript/api/excel/excel.shape#excel-excel-shape-copyto-member(1))|Copies and pastes a `Shape` object.|
||[placement](/javascript/api/excel/excel.shape#excel-excel-shape-placement-member)|Represents how the object is attached to the cells below it.|
|[Slicer](/javascript/api/excel/excel.slicer)|[caption](/javascript/api/excel/excel.slicer#excel-excel-slicer-caption-member)|Represents the caption of the slicer.|
||[clearFilters()](/javascript/api/excel/excel.slicer#excel-excel-slicer-clearfilters-member(1))|Clears all the filters currently applied on the slicer.|
||[delete()](/javascript/api/excel/excel.slicer#excel-excel-slicer-delete-member(1))|Deletes the slicer.|
||[getSelectedItems()](/javascript/api/excel/excel.slicer#excel-excel-slicer-getselecteditems-member(1))|Returns an array of selected items' keys.|
||[height](/javascript/api/excel/excel.slicer#excel-excel-slicer-height-member)|Represents the height, in points, of the slicer.|
||[id](/javascript/api/excel/excel.slicer#excel-excel-slicer-id-member)|Represents the unique ID of the slicer.|
||[isFilterCleared](/javascript/api/excel/excel.slicer#excel-excel-slicer-isfiltercleared-member)|Value is `true` if all filters currently applied on the slicer are cleared.|
||[left](/javascript/api/excel/excel.slicer#excel-excel-slicer-left-member)|Represents the distance, in points, from the left side of the slicer to the left of the worksheet.|
||[name](/javascript/api/excel/excel.slicer#excel-excel-slicer-name-member)|Represents the name of the slicer.|
||[selectItems(items?: string[])](/javascript/api/excel/excel.slicer#excel-excel-slicer-selectitems-member(1))|Selects slicer items based on their keys.|
||[slicerItems](/javascript/api/excel/excel.slicer#excel-excel-slicer-sliceritems-member)|Represents the collection of slicer items that are part of the slicer.|
||[sortBy](/javascript/api/excel/excel.slicer#excel-excel-slicer-sortby-member)|Represents the sort order of the items in the slicer.|
||[style](/javascript/api/excel/excel.slicer#excel-excel-slicer-style-member)|Constant value that represents the slicer style.|
||[top](/javascript/api/excel/excel.slicer#excel-excel-slicer-top-member)|Represents the distance, in points, from the top edge of the slicer to the top of the worksheet.|
||[width](/javascript/api/excel/excel.slicer#excel-excel-slicer-width-member)|Represents the width, in points, of the slicer.|
||[worksheet](/javascript/api/excel/excel.slicer#excel-excel-slicer-worksheet-member)|Represents the worksheet containing the slicer.|
|[SlicerCollection](/javascript/api/excel/excel.slicercollection)|[add(slicerSource: string \| PivotTable \| Table, sourceField: string \| PivotField \| number \| TableColumn, slicerDestination?: string \| Worksheet)](/javascript/api/excel/excel.slicercollection#excel-excel-slicercollection-add-member(1))|Adds a new slicer to the workbook.|
||[getCount()](/javascript/api/excel/excel.slicercollection#excel-excel-slicercollection-getcount-member(1))|Returns the number of slicers in the collection.|
||[getItem(key: string)](/javascript/api/excel/excel.slicercollection#excel-excel-slicercollection-getitem-member(1))|Gets a slicer object using its name or ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.slicercollection#excel-excel-slicercollection-getitemat-member(1))|Gets a slicer based on its position in the collection.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.slicercollection#excel-excel-slicercollection-getitemornullobject-member(1))|Gets a slicer using its name or ID.|
||[items](/javascript/api/excel/excel.slicercollection#excel-excel-slicercollection-items-member)|Gets the loaded child items in this collection.|
|[SlicerItem](/javascript/api/excel/excel.sliceritem)|[hasData](/javascript/api/excel/excel.sliceritem#excel-excel-sliceritem-hasdata-member)|Value is `true` if the slicer item has data.|
||[isSelected](/javascript/api/excel/excel.sliceritem#excel-excel-sliceritem-isselected-member)|Value is `true` if the slicer item is selected.|
||[key](/javascript/api/excel/excel.sliceritem#excel-excel-sliceritem-key-member)|Represents the unique value representing the slicer item.|
||[name](/javascript/api/excel/excel.sliceritem#excel-excel-sliceritem-name-member)|Represents the title displayed in the Excel UI.|
|[SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection)|[getCount()](/javascript/api/excel/excel.sliceritemcollection#excel-excel-sliceritemcollection-getcount-member(1))|Returns the number of slicer items in the slicer.|
||[getItem(key: string)](/javascript/api/excel/excel.sliceritemcollection#excel-excel-sliceritemcollection-getitem-member(1))|Gets a slicer item object using its key or name.|
||[getItemAt(index: number)](/javascript/api/excel/excel.sliceritemcollection#excel-excel-sliceritemcollection-getitemat-member(1))|Gets a slicer item based on its position in the collection.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.sliceritemcollection#excel-excel-sliceritemcollection-getitemornullobject-member(1))|Gets a slicer item using its key or name.|
||[items](/javascript/api/excel/excel.sliceritemcollection#excel-excel-sliceritemcollection-items-member)|Gets the loaded child items in this collection.|
|[SlicerStyle](/javascript/api/excel/excel.slicerstyle)|[delete()](/javascript/api/excel/excel.slicerstyle#excel-excel-slicerstyle-delete-member(1))|Deletes the slicer style.|
||[duplicate()](/javascript/api/excel/excel.slicerstyle#excel-excel-slicerstyle-duplicate-member(1))|Creates a duplicate of this slicer style with copies of all the style elements.|
||[name](/javascript/api/excel/excel.slicerstyle#excel-excel-slicerstyle-name-member)|Gets the name of the slicer style.|
||[readOnly](/javascript/api/excel/excel.slicerstyle#excel-excel-slicerstyle-readonly-member)|Specifies if this `SlicerStyle` object is read-only.|
|[SlicerStyleCollection](/javascript/api/excel/excel.slicerstylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.slicerstylecollection#excel-excel-slicerstylecollection-add-member(1))|Creates a blank slicer style with the specified name.|
||[getCount()](/javascript/api/excel/excel.slicerstylecollection#excel-excel-slicerstylecollection-getcount-member(1))|Gets the number of slicer styles in the collection.|
||[getDefault()](/javascript/api/excel/excel.slicerstylecollection#excel-excel-slicerstylecollection-getdefault-member(1))|Gets the default `SlicerStyle` for the parent object's scope.|
||[getItem(name: string)](/javascript/api/excel/excel.slicerstylecollection#excel-excel-slicerstylecollection-getitem-member(1))|Gets a `SlicerStyle` by name.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.slicerstylecollection#excel-excel-slicerstylecollection-getitemornullobject-member(1))|Gets a `SlicerStyle` by name.|
||[items](/javascript/api/excel/excel.slicerstylecollection#excel-excel-slicerstylecollection-items-member)|Gets the loaded child items in this collection.|
||[setDefault(newDefaultStyle: SlicerStyle \| string)](/javascript/api/excel/excel.slicerstylecollection#excel-excel-slicerstylecollection-setdefault-member(1))|Sets the default slicer style for use in the parent object's scope.|
|[TableStyle](/javascript/api/excel/excel.tablestyle)|[delete()](/javascript/api/excel/excel.tablestyle#excel-excel-tablestyle-delete-member(1))|Deletes the table style.|
||[duplicate()](/javascript/api/excel/excel.tablestyle#excel-excel-tablestyle-duplicate-member(1))|Creates a duplicate of this table style with copies of all the style elements.|
||[name](/javascript/api/excel/excel.tablestyle#excel-excel-tablestyle-name-member)|Gets the name of the table style.|
||[readOnly](/javascript/api/excel/excel.tablestyle#excel-excel-tablestyle-readonly-member)|Specifies if this `TableStyle` object is read-only.|
|[TableStyleCollection](/javascript/api/excel/excel.tablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.tablestylecollection#excel-excel-tablestylecollection-add-member(1))|Creates a blank `TableStyle` with the specified name.|
||[getCount()](/javascript/api/excel/excel.tablestylecollection#excel-excel-tablestylecollection-getcount-member(1))|Gets the number of table styles in the collection.|
||[getDefault()](/javascript/api/excel/excel.tablestylecollection#excel-excel-tablestylecollection-getdefault-member(1))|Gets the default table style for the parent object's scope.|
||[getItem(name: string)](/javascript/api/excel/excel.tablestylecollection#excel-excel-tablestylecollection-getitem-member(1))|Gets a `TableStyle` by name.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.tablestylecollection#excel-excel-tablestylecollection-getitemornullobject-member(1))|Gets a `TableStyle` by name.|
||[items](/javascript/api/excel/excel.tablestylecollection#excel-excel-tablestylecollection-items-member)|Gets the loaded child items in this collection.|
||[setDefault(newDefaultStyle: TableStyle \| string)](/javascript/api/excel/excel.tablestylecollection#excel-excel-tablestylecollection-setdefault-member(1))|Sets the default table style for use in the parent object's scope.|
|[TimelineStyle](/javascript/api/excel/excel.timelinestyle)|[delete()](/javascript/api/excel/excel.timelinestyle#excel-excel-timelinestyle-delete-member(1))|Deletes the table style.|
||[duplicate()](/javascript/api/excel/excel.timelinestyle#excel-excel-timelinestyle-duplicate-member(1))|Creates a duplicate of this timeline style with copies of all the style elements.|
||[name](/javascript/api/excel/excel.timelinestyle#excel-excel-timelinestyle-name-member)|Gets the name of the timeline style.|
||[readOnly](/javascript/api/excel/excel.timelinestyle#excel-excel-timelinestyle-readonly-member)|Specifies if this `TimelineStyle` object is read-only.|
|[TimelineStyleCollection](/javascript/api/excel/excel.timelinestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.timelinestylecollection#excel-excel-timelinestylecollection-add-member(1))|Creates a blank `TimelineStyle` with the specified name.|
||[getCount()](/javascript/api/excel/excel.timelinestylecollection#excel-excel-timelinestylecollection-getcount-member(1))|Gets the number of timeline styles in the collection.|
||[getDefault()](/javascript/api/excel/excel.timelinestylecollection#excel-excel-timelinestylecollection-getdefault-member(1))|Gets the default timeline style for the parent object's scope.|
||[getItem(name: string)](/javascript/api/excel/excel.timelinestylecollection#excel-excel-timelinestylecollection-getitem-member(1))|Gets a `TimelineStyle` by name.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.timelinestylecollection#excel-excel-timelinestylecollection-getitemornullobject-member(1))|Gets a `TimelineStyle` by name.|
||[items](/javascript/api/excel/excel.timelinestylecollection#excel-excel-timelinestylecollection-items-member)|Gets the loaded child items in this collection.|
||[setDefault(newDefaultStyle: TimelineStyle \| string)](/javascript/api/excel/excel.timelinestylecollection#excel-excel-timelinestylecollection-setdefault-member(1))|Sets the default timeline style for use in the parent object's scope.|
|[Workbook](/javascript/api/excel/excel.workbook)|[comments](/javascript/api/excel/excel.workbook#excel-excel-workbook-comments-member)|Represents a collection of comments associated with the workbook.|
||[getActiveSlicer()](/javascript/api/excel/excel.workbook#excel-excel-workbook-getactiveslicer-member(1))|Gets the currently active slicer in the workbook.|
||[getActiveSlicerOrNullObject()](/javascript/api/excel/excel.workbook#excel-excel-workbook-getactiveslicerornullobject-member(1))|Gets the currently active slicer in the workbook.|
||[pivotTableStyles](/javascript/api/excel/excel.workbook#excel-excel-workbook-pivottablestyles-member)|Represents a collection of PivotTableStyles associated with the workbook.|
||[slicerStyles](/javascript/api/excel/excel.workbook#excel-excel-workbook-slicerstyles-member)|Represents a collection of SlicerStyles associated with the workbook.|
||[slicers](/javascript/api/excel/excel.workbook#excel-excel-workbook-slicers-member)|Represents a collection of slicers associated with the workbook.|
||[tableStyles](/javascript/api/excel/excel.workbook#excel-excel-workbook-tablestyles-member)|Represents a collection of TableStyles associated with the workbook.|
||[timelineStyles](/javascript/api/excel/excel.workbook#excel-excel-workbook-timelinestyles-member)|Represents a collection of TimelineStyles associated with the workbook.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[comments](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-comments-member)|Returns a collection of all the Comments objects on the worksheet.|
||[onColumnSorted](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-oncolumnsorted-member)|Occurs when one or more columns have been sorted.|
||[onRowSorted](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onrowsorted-member)|Occurs when one or more rows have been sorted.|
||[onSingleClicked](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onsingleclicked-member)|Occurs when a left-clicked/tapped action happens in the worksheet.|
||[showOutlineLevels(rowLevels: number, columnLevels: number)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-showoutlinelevels-member(1))|Shows row or column groups by their outline levels.|
||[slicers](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-slicers-member)|Returns a collection of slicers that are part of the worksheet.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onColumnSorted](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-oncolumnsorted-member)|Occurs when one or more columns have been sorted.|
||[onRowSorted](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onrowsorted-member)|Occurs when one or more rows have been sorted.|
||[onSingleClicked](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onsingleclicked-member)|Occurs when left-clicked/tapped operation happens in the worksheet collection.|
|[WorksheetColumnSortedEventArgs](/javascript/api/excel/excel.worksheetcolumnsortedeventargs)|[address](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#excel-excel-worksheetcolumnsortedeventargs-address-member)|Gets the range address that represents the sorted areas of a specific worksheet.|
||[source](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#excel-excel-worksheetcolumnsortedeventargs-source-member)|Gets the source of the event.|
||[type](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#excel-excel-worksheetcolumnsortedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#excel-excel-worksheetcolumnsortedeventargs-worksheetid-member)|Gets the ID of the worksheet where the sorting happened.|
|[WorksheetRowSortedEventArgs](/javascript/api/excel/excel.worksheetrowsortedeventargs)|[address](/javascript/api/excel/excel.worksheetrowsortedeventargs#excel-excel-worksheetrowsortedeventargs-address-member)|Gets the range address that represents the sorted areas of a specific worksheet.|
||[source](/javascript/api/excel/excel.worksheetrowsortedeventargs#excel-excel-worksheetrowsortedeventargs-source-member)|Gets the source of the event.|
||[type](/javascript/api/excel/excel.worksheetrowsortedeventargs#excel-excel-worksheetrowsortedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowsortedeventargs#excel-excel-worksheetrowsortedeventargs-worksheetid-member)|Gets the ID of the worksheet where the sorting happened.|
|[WorksheetSingleClickedEventArgs](/javascript/api/excel/excel.worksheetsingleclickedeventargs)|[address](/javascript/api/excel/excel.worksheetsingleclickedeventargs#excel-excel-worksheetsingleclickedeventargs-address-member)|Gets the address that represents the cell which was left-clicked/tapped for a specific worksheet.|
||[offsetX](/javascript/api/excel/excel.worksheetsingleclickedeventargs#excel-excel-worksheetsingleclickedeventargs-offsetx-member)|The distance, in points, from the left-clicked/tapped point to the left (or right for right-to-left languages) gridline edge of the left-clicked/tapped cell.|
||[offsetY](/javascript/api/excel/excel.worksheetsingleclickedeventargs#excel-excel-worksheetsingleclickedeventargs-offsety-member)|The distance, in points, from the left-clicked/tapped point to the top gridline edge of the left-clicked/tapped cell.|
||[type](/javascript/api/excel/excel.worksheetsingleclickedeventargs#excel-excel-worksheetsingleclickedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetsingleclickedeventargs#excel-excel-worksheetsingleclickedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the cell was left-clicked/tapped.|

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel?view=excel-js-1.10&preserve-view=true)
- [Excel JavaScript API requirement sets](excel-api-requirement-sets.md)