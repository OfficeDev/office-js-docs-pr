---
title: Excel JavaScript preview APIs
description: 'Details about upcoming Excel JavaScript APIs'
ms.date: 07/11/2019
ms.prod: excel
localization_priority: Normal
---

# Excel JavaScript preview APIs

New Excel JavaScript APIs are first introduced in "preview" and later become part of a specific, numbered requirement set after sufficient testing occurs and user feedback is acquired.

The first table provides a concise summary of the APIs, while the subsequent table gives a detailed list.

> [!NOTE]
> Preview APIs are subject to change and are not intended for use in a production environment. We recommend that you try them out in test and development environments only. Do not use preview APIs in a production environment or within business-critical documents.
>
> To use preview APIs, you must reference the **beta** library on the CDN (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) and you may also need to join the Office Insider program to get a recent Office build.

| Feature area | Description | Relevant objects |
|:--- |:--- |:--- |
| [Slicer](../../excel/excel-add-ins-pivottables.md#slicers-preview) | Insert and configure slicers to tables and PivotTables. | [Slicer](/javascript/api/excel/excel.slicer) |
| [Comments](../../excel/excel-add-ins-workbooks.md#comments-preview) | Add, edit, and delete comments. | [Comment](/javascript/api/excel/excel.comment), [CommentCollection](/javascript/api/excel/excel.commentcollection) |
| Workbook [Save](../../excel/excel-add-ins-workbooks.md#save-the-workbook-preview) and [Close](../../excel/excel-add-ins-workbooks.md#close-the-workbook-preview) | Save and close workbooks.  | [Workbook](/javascript/api/excel/excel.workbook) |
| [Insert Workbook](../../excel/excel-add-ins-workbooks.md#insert-a-copy-of-an-existing-workbook-into-the-current-one-preview) | Insert one workbook into another.  | [Workbook](/javascript/api/excel/excel.worksheetcollection) |

## API list

| Class | Fields | Description |
|:---|:---|:---|
|[Comment](/javascript/api/excel/excel.comment)|[content](/javascript/api/excel/excel.comment#content)|Gets or sets the comment's content. The string is plain text.|
||[delete()](/javascript/api/excel/excel.comment#delete--)|Deletes the comment thread.|
||[getLocation()](/javascript/api/excel/excel.comment#getlocation--)|Gets the cell where this comment is located.|
||[authorEmail](/javascript/api/excel/excel.comment#authoremail)|Gets the email of the comment's author.|
||[authorName](/javascript/api/excel/excel.comment#authorname)|Gets the name of the comment's author.|
||[creationDate](/javascript/api/excel/excel.comment#creationdate)|Gets the creation time of the comment. Returns null if the comment was converted from a note, since the comment does not have a creation date.|
||[id](/javascript/api/excel/excel.comment#id)|Represents the comment identifier. Read-only.|
||[replies](/javascript/api/excel/excel.comment#replies)|Represents a collection of reply objects associated with the comment. Read-only.|
||[set(properties: Excel.Comment)](/javascript/api/excel/excel.comment#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.CommentUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.comment#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add(content: string, cellAddress: Range \| string, contentType?: "Plain")](/javascript/api/excel/excel.commentcollection#add-content--celladdress--contenttype-)|Creates a new comment (comment thread) with the given content on the given cell. An `InvalidArgument` error is thrown if the provided range is larger than one cell.|
||[add(content: string, cellAddress: Range \| string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentcollection#add-content--celladdress--contenttype-)|Creates a new comment (comment thread) with the given content on the given cell. An `InvalidArgument` error is thrown if the provided range is larger than one cell.|
||[getCount()](/javascript/api/excel/excel.commentcollection#getcount--)|Gets the number of comments in the collection.|
||[getItem(commentId: string)](/javascript/api/excel/excel.commentcollection#getitem-commentid-)|Gets a comment from the collection based on its ID. Read-only.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentcollection#getitemat-index-)|Gets a comment from the collection based on its position.|
||[getItemByCell(cellAddress: Range \| string)](/javascript/api/excel/excel.commentcollection#getitembycell-celladdress-)|Gets the comment from the specified cell.|
||[getItemByReplyId(replyId: string)](/javascript/api/excel/excel.commentcollection#getitembyreplyid-replyid-)|Gets a comment related to its reply ID in the collection.|
||[items](/javascript/api/excel/excel.commentcollection#items)|Gets the loaded child items in this collection.|
|[CommentCollectionData](/javascript/api/excel/excel.commentcollectiondata)|[items](/javascript/api/excel/excel.commentcollectiondata#items)||
|[CommentCollectionLoadOptions](/javascript/api/excel/excel.commentcollectionloadoptions)|[$all](/javascript/api/excel/excel.commentcollectionloadoptions#$all)||
||[authorEmail](/javascript/api/excel/excel.commentcollectionloadoptions#authoremail)|For EACH ITEM in the collection: Gets the email of the comment's author.|
||[authorName](/javascript/api/excel/excel.commentcollectionloadoptions#authorname)|For EACH ITEM in the collection: Gets the name of the comment's author.|
||[content](/javascript/api/excel/excel.commentcollectionloadoptions#content)|For EACH ITEM in the collection: Gets or sets the comment's content. The string is plain text.|
||[creationDate](/javascript/api/excel/excel.commentcollectionloadoptions#creationdate)|For EACH ITEM in the collection: Gets the creation time of the comment. Returns null if the comment was converted from a note, since the comment does not have a creation date.|
||[id](/javascript/api/excel/excel.commentcollectionloadoptions#id)|For EACH ITEM in the collection: Represents the comment identifier. Read-only.|
|[CommentCollectionUpdateData](/javascript/api/excel/excel.commentcollectionupdatedata)|[items](/javascript/api/excel/excel.commentcollectionupdatedata#items)||
|[CommentData](/javascript/api/excel/excel.commentdata)|[authorEmail](/javascript/api/excel/excel.commentdata#authoremail)|Gets the email of the comment's author.|
||[authorName](/javascript/api/excel/excel.commentdata#authorname)|Gets the name of the comment's author.|
||[content](/javascript/api/excel/excel.commentdata#content)|Gets or sets the comment's content. The string is plain text.|
||[creationDate](/javascript/api/excel/excel.commentdata#creationdate)|Gets the creation time of the comment. Returns null if the comment was converted from a note, since the comment does not have a creation date.|
||[id](/javascript/api/excel/excel.commentdata#id)|Represents the comment identifier. Read-only.|
||[replies](/javascript/api/excel/excel.commentdata#replies)|Represents a collection of reply objects associated with the comment. Read-only.|
|[CommentLoadOptions](/javascript/api/excel/excel.commentloadoptions)|[$all](/javascript/api/excel/excel.commentloadoptions#$all)||
||[authorEmail](/javascript/api/excel/excel.commentloadoptions#authoremail)|Gets the email of the comment's author.|
||[authorName](/javascript/api/excel/excel.commentloadoptions#authorname)|Gets the name of the comment's author.|
||[content](/javascript/api/excel/excel.commentloadoptions#content)|Gets or sets the comment's content. The string is plain text.|
||[creationDate](/javascript/api/excel/excel.commentloadoptions#creationdate)|Gets the creation time of the comment. Returns null if the comment was converted from a note, since the comment does not have a creation date.|
||[id](/javascript/api/excel/excel.commentloadoptions#id)|Represents the comment identifier. Read-only.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[content](/javascript/api/excel/excel.commentreply#content)|Gets or sets the comment reply's content. The string is plain text.|
||[delete()](/javascript/api/excel/excel.commentreply#delete--)|Deletes the comment reply.|
||[getLocation()](/javascript/api/excel/excel.commentreply#getlocation--)|Gets the cell where this comment reply is located.|
||[getParentComment()](/javascript/api/excel/excel.commentreply#getparentcomment--)|Gets the parent comment of this reply.|
||[authorEmail](/javascript/api/excel/excel.commentreply#authoremail)|Gets the email of the comment reply's author.|
||[authorName](/javascript/api/excel/excel.commentreply#authorname)|Gets the name of the comment reply's author.|
||[creationDate](/javascript/api/excel/excel.commentreply#creationdate)|Gets the creation time of the comment reply.|
||[id](/javascript/api/excel/excel.commentreply#id)|Represents the comment reply identifier. Read-only.|
||[set(properties: Excel.CommentReply)](/javascript/api/excel/excel.commentreply#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.CommentReplyUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.commentreply#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add(content: string, contentType?: "Plain")](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|Creates a comment reply for comment.|
||[add(content: string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|Creates a comment reply for comment.|
||[getCount()](/javascript/api/excel/excel.commentreplycollection#getcount--)|Gets the number of comment replies in the collection.|
||[getItem(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#getitem-commentreplyid-)|Returns a comment reply identified by its ID. Read-only.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentreplycollection#getitemat-index-)|Gets a comment reply based on its position in the collection.|
||[items](/javascript/api/excel/excel.commentreplycollection#items)|Gets the loaded child items in this collection.|
|[CommentReplyCollectionData](/javascript/api/excel/excel.commentreplycollectiondata)|[items](/javascript/api/excel/excel.commentreplycollectiondata#items)||
|[CommentReplyCollectionLoadOptions](/javascript/api/excel/excel.commentreplycollectionloadoptions)|[$all](/javascript/api/excel/excel.commentreplycollectionloadoptions#$all)||
||[authorEmail](/javascript/api/excel/excel.commentreplycollectionloadoptions#authoremail)|For EACH ITEM in the collection: Gets the email of the comment reply's author.|
||[authorName](/javascript/api/excel/excel.commentreplycollectionloadoptions#authorname)|For EACH ITEM in the collection: Gets the name of the comment reply's author.|
||[content](/javascript/api/excel/excel.commentreplycollectionloadoptions#content)|For EACH ITEM in the collection: Gets or sets the comment reply's content. The string is plain text.|
||[creationDate](/javascript/api/excel/excel.commentreplycollectionloadoptions#creationdate)|For EACH ITEM in the collection: Gets the creation time of the comment reply.|
||[id](/javascript/api/excel/excel.commentreplycollectionloadoptions#id)|For EACH ITEM in the collection: Represents the comment reply identifier. Read-only.|
|[CommentReplyCollectionUpdateData](/javascript/api/excel/excel.commentreplycollectionupdatedata)|[items](/javascript/api/excel/excel.commentreplycollectionupdatedata#items)||
|[CommentReplyData](/javascript/api/excel/excel.commentreplydata)|[authorEmail](/javascript/api/excel/excel.commentreplydata#authoremail)|Gets the email of the comment reply's author.|
||[authorName](/javascript/api/excel/excel.commentreplydata#authorname)|Gets the name of the comment reply's author.|
||[content](/javascript/api/excel/excel.commentreplydata#content)|Gets or sets the comment reply's content. The string is plain text.|
||[creationDate](/javascript/api/excel/excel.commentreplydata#creationdate)|Gets the creation time of the comment reply.|
||[id](/javascript/api/excel/excel.commentreplydata#id)|Represents the comment reply identifier. Read-only.|
|[CommentReplyLoadOptions](/javascript/api/excel/excel.commentreplyloadoptions)|[$all](/javascript/api/excel/excel.commentreplyloadoptions#$all)||
||[authorEmail](/javascript/api/excel/excel.commentreplyloadoptions#authoremail)|Gets the email of the comment reply's author.|
||[authorName](/javascript/api/excel/excel.commentreplyloadoptions#authorname)|Gets the name of the comment reply's author.|
||[content](/javascript/api/excel/excel.commentreplyloadoptions#content)|Gets or sets the comment reply's content. The string is plain text.|
||[creationDate](/javascript/api/excel/excel.commentreplyloadoptions#creationdate)|Gets the creation time of the comment reply.|
||[id](/javascript/api/excel/excel.commentreplyloadoptions#id)|Represents the comment reply identifier. Read-only.|
|[CommentReplyUpdateData](/javascript/api/excel/excel.commentreplyupdatedata)|[content](/javascript/api/excel/excel.commentreplyupdatedata#content)|Gets or sets the comment reply's content. The string is plain text.|
|[CommentUpdateData](/javascript/api/excel/excel.commentupdatedata)|[content](/javascript/api/excel/excel.commentupdatedata#content)|Gets or sets the comment's content. The string is plain text.|
|[GroupShapeCollectionLoadOptions](/javascript/api/excel/excel.groupshapecollectionloadoptions)|[placement](/javascript/api/excel/excel.groupshapecollectionloadoptions#placement)|For EACH ITEM in the collection: Represents how the object is attached to the cells below it.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[enableFieldList](/javascript/api/excel/excel.pivotlayout#enablefieldlist)|Specifies whether the field list can be shown in the UI.|
||[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|Gets a unique cell in the PivotTable based on a data hierarchy and the row and column items of their respective hierarchies. The returned cell is the intersection of the given row and column that contains the data from the given hierarchy. This method is the inverse of calling getPivotItems and getDataHierarchy on a particular cell.|
|[PivotLayoutData](/javascript/api/excel/excel.pivotlayoutdata)|[enableFieldList](/javascript/api/excel/excel.pivotlayoutdata#enablefieldlist)|Specifies whether the field list can be shown in the UI.|
|[PivotLayoutLoadOptions](/javascript/api/excel/excel.pivotlayoutloadoptions)|[enableFieldList](/javascript/api/excel/excel.pivotlayoutloadoptions#enablefieldlist)|Specifies whether the field list can be shown in the UI.|
|[PivotLayoutUpdateData](/javascript/api/excel/excel.pivotlayoutupdatedata)|[enableFieldList](/javascript/api/excel/excel.pivotlayoutupdatedata#enablefieldlist)|Specifies whether the field list can be shown in the UI.|
|[PivotTableStyle](/javascript/api/excel/excel.pivottablestyle)|[delete()](/javascript/api/excel/excel.pivottablestyle#delete--)|Deletes the PivotTableStyle.|
||[duplicate()](/javascript/api/excel/excel.pivottablestyle#duplicate--)|Creates a duplicate of this PivotTableStyle with copies of all the style elements.|
||[name](/javascript/api/excel/excel.pivottablestyle#name)|Gets the name of the PivotTableStyle.|
||[readOnly](/javascript/api/excel/excel.pivottablestyle#readonly)|Specifies if this PivotTableStyle object is read-only. Read-only.|
||[set(properties: Excel.PivotTableStyle)](/javascript/api/excel/excel.pivottablestyle#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.PivotTableStyleUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.pivottablestyle#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
|[PivotTableStyleCollection](/javascript/api/excel/excel.pivottablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.pivottablestylecollection#add-name--makeuniquename-)|Creates a blank PivotTableStyle with the specified name.|
||[getCount()](/javascript/api/excel/excel.pivottablestylecollection#getcount--)|Gets the number of PivotTable styles in the collection.|
||[getDefault()](/javascript/api/excel/excel.pivottablestylecollection#getdefault--)|Gets the default PivotTableStyle for the parent object's scope.|
||[getItem(name: string)](/javascript/api/excel/excel.pivottablestylecollection#getitem-name-)|Gets a PivotTableStyle by name.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablestylecollection#getitemornullobject-name-)|Gets a PivotTableStyle by name. If the PivotTableStyle does not exist, will return a null object.|
||[items](/javascript/api/excel/excel.pivottablestylecollection#items)|Gets the loaded child items in this collection.|
||[setDefault(newDefaultStyle: PivotTableStyle \| string)](/javascript/api/excel/excel.pivottablestylecollection#setdefault-newdefaultstyle-)|Sets the default PivotTableStyle for use in the parent object's scope.|
|[PivotTableStyleCollectionData](/javascript/api/excel/excel.pivottablestylecollectiondata)|[items](/javascript/api/excel/excel.pivottablestylecollectiondata#items)||
|[PivotTableStyleCollectionLoadOptions](/javascript/api/excel/excel.pivottablestylecollectionloadoptions)|[$all](/javascript/api/excel/excel.pivottablestylecollectionloadoptions#$all)||
||[name](/javascript/api/excel/excel.pivottablestylecollectionloadoptions#name)|For EACH ITEM in the collection: Gets the name of the PivotTableStyle.|
||[readOnly](/javascript/api/excel/excel.pivottablestylecollectionloadoptions#readonly)|For EACH ITEM in the collection: Specifies if this PivotTableStyle object is read-only. Read-only.|
|[PivotTableStyleCollectionUpdateData](/javascript/api/excel/excel.pivottablestylecollectionupdatedata)|[items](/javascript/api/excel/excel.pivottablestylecollectionupdatedata#items)||
|[PivotTableStyleData](/javascript/api/excel/excel.pivottablestyledata)|[name](/javascript/api/excel/excel.pivottablestyledata#name)|Gets the name of the PivotTableStyle.|
||[readOnly](/javascript/api/excel/excel.pivottablestyledata#readonly)|Specifies if this PivotTableStyle object is read-only. Read-only.|
|[PivotTableStyleLoadOptions](/javascript/api/excel/excel.pivottablestyleloadoptions)|[$all](/javascript/api/excel/excel.pivottablestyleloadoptions#$all)||
||[name](/javascript/api/excel/excel.pivottablestyleloadoptions#name)|Gets the name of the PivotTableStyle.|
||[readOnly](/javascript/api/excel/excel.pivottablestyleloadoptions#readonly)|Specifies if this PivotTableStyle object is read-only. Read-only.|
|[PivotTableStyleUpdateData](/javascript/api/excel/excel.pivottablestyleupdatedata)|[name](/javascript/api/excel/excel.pivottablestyleupdatedata#name)|Gets the name of the PivotTableStyle.|
|[Range](/javascript/api/excel/excel.range)|[getSpillParent()](/javascript/api/excel/excel.range#getspillparent--)|Gets the range object containing the anchor cell for a cell getting spilled into. Fails if applied to a range with more than one cell. Read-only.|
||[getSpillParentOrNullObject()](/javascript/api/excel/excel.range#getspillparentornullobject--)|Gets the range object containing the anchor cell for a cell getting spilled into. Read-only.|
||[getSpillingToRange()](/javascript/api/excel/excel.range#getspillingtorange--)|Gets the range object containing the spill range when called on an anchor cell. Fails if applied to a range with more than one cell. Read-only.|
||[getSpillingToRangeOrNullObject()](/javascript/api/excel/excel.range#getspillingtorangeornullobject--)|Gets the range object containing the spill range when called on an anchor cell. Read-only.|
||[hasSpill](/javascript/api/excel/excel.range#hasspill)|Represents if all cells have a spill border.|
||[height](/javascript/api/excel/excel.range#height)|Returns the distance in points, for 100% zoom, from top edge of the range to bottom edge of the range. Read-only.|
||[left](/javascript/api/excel/excel.range#left)|Returns the distance in points, for 100% zoom, from left edge of the worksheet to left edge of the range. Read-only.|
||[savedAsArray](/javascript/api/excel/excel.range#savedasarray)|Represents if ALL the cells would be saved as an array formula.|
||[top](/javascript/api/excel/excel.range#top)|Returns the distance in points, for 100% zoom, from top edge of the worksheet to top edge of the range. Read-only.|
||[width](/javascript/api/excel/excel.range#width)|Returns the distance in points, for 100% zoom, from left edge of the range to right edge of the range. Read-only.|
|[RangeCollectionLoadOptions](/javascript/api/excel/excel.rangecollectionloadoptions)|[hasSpill](/javascript/api/excel/excel.rangecollectionloadoptions#hasspill)|For EACH ITEM in the collection: Represents if all cells have a spill border.|
||[height](/javascript/api/excel/excel.rangecollectionloadoptions#height)|For EACH ITEM in the collection: Returns the distance in points, for 100% zoom, from top edge of the range to bottom edge of the range. Read-only.|
||[left](/javascript/api/excel/excel.rangecollectionloadoptions#left)|For EACH ITEM in the collection: Returns the distance in points, for 100% zoom, from left edge of the worksheet to left edge of the range. Read-only.|
||[savedAsArray](/javascript/api/excel/excel.rangecollectionloadoptions#savedasarray)|For EACH ITEM in the collection: Represents if ALL the cells would be saved as an array formula.|
||[top](/javascript/api/excel/excel.rangecollectionloadoptions#top)|For EACH ITEM in the collection: Returns the distance in points, for 100% zoom, from top edge of the worksheet to top edge of the range. Read-only.|
||[width](/javascript/api/excel/excel.rangecollectionloadoptions#width)|For EACH ITEM in the collection: Returns the distance in points, for 100% zoom, from left edge of the range to right edge of the range. Read-only.|
|[RangeData](/javascript/api/excel/excel.rangedata)|[hasSpill](/javascript/api/excel/excel.rangedata#hasspill)|Represents if all cells have a spill border.|
||[height](/javascript/api/excel/excel.rangedata#height)|Returns the distance in points, for 100% zoom, from top edge of the range to bottom edge of the range. Read-only.|
||[left](/javascript/api/excel/excel.rangedata#left)|Returns the distance in points, for 100% zoom, from left edge of the worksheet to left edge of the range. Read-only.|
||[savedAsArray](/javascript/api/excel/excel.rangedata#savedasarray)|Represents if ALL the cells would be saved as an array formula.|
||[top](/javascript/api/excel/excel.rangedata#top)|Returns the distance in points, for 100% zoom, from top edge of the worksheet to top edge of the range. Read-only.|
||[width](/javascript/api/excel/excel.rangedata#width)|Returns the distance in points, for 100% zoom, from left edge of the range to right edge of the range. Read-only.|
|[RangeLoadOptions](/javascript/api/excel/excel.rangeloadoptions)|[hasSpill](/javascript/api/excel/excel.rangeloadoptions#hasspill)|Represents if all cells have a spill border.|
||[height](/javascript/api/excel/excel.rangeloadoptions#height)|Returns the distance in points, for 100% zoom, from top edge of the range to bottom edge of the range. Read-only.|
||[left](/javascript/api/excel/excel.rangeloadoptions#left)|Returns the distance in points, for 100% zoom, from left edge of the worksheet to left edge of the range. Read-only.|
||[savedAsArray](/javascript/api/excel/excel.rangeloadoptions#savedasarray)|Represents if ALL the cells would be saved as an array formula.|
||[top](/javascript/api/excel/excel.rangeloadoptions#top)|Returns the distance in points, for 100% zoom, from top edge of the worksheet to top edge of the range. Read-only.|
||[width](/javascript/api/excel/excel.rangeloadoptions#width)|Returns the distance in points, for 100% zoom, from left edge of the range to right edge of the range. Read-only.|
|[Shape](/javascript/api/excel/excel.shape)|[copyTo(destinationSheet?: Worksheet \| string)](/javascript/api/excel/excel.shape#copyto-destinationsheet-)|Copies and pastes a Shape object.|
||[placement](/javascript/api/excel/excel.shape#placement)|Represents how the object is attached to the cells below it.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|Creates a scalable vector graphic (SVG) from an XML string and adds it to the worksheet. Returns a Shape object that represents the new image.|
|[ShapeCollectionLoadOptions](/javascript/api/excel/excel.shapecollectionloadoptions)|[placement](/javascript/api/excel/excel.shapecollectionloadoptions#placement)|For EACH ITEM in the collection: Represents how the object is attached to the cells below it.|
|[ShapeData](/javascript/api/excel/excel.shapedata)|[placement](/javascript/api/excel/excel.shapedata#placement)|Represents how the object is attached to the cells below it.|
|[ShapeLoadOptions](/javascript/api/excel/excel.shapeloadoptions)|[placement](/javascript/api/excel/excel.shapeloadoptions#placement)|Represents how the object is attached to the cells below it.|
|[ShapeUpdateData](/javascript/api/excel/excel.shapeupdatedata)|[placement](/javascript/api/excel/excel.shapeupdatedata#placement)|Represents how the object is attached to the cells below it.|
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
||[selectItems(items?: string[])](/javascript/api/excel/excel.slicer#selectitems-items-)|Select slicer items based on their keys. Previous selection will be cleared.|
||[set(properties: Excel.Slicer)](/javascript/api/excel/excel.slicer#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.SlicerUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.slicer#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
||[sortBy](/javascript/api/excel/excel.slicer#sortby)|Represents the sort order of the items in the slicer. Possible values are: DataSourceOrder, Ascending, Descending.|
||[style](/javascript/api/excel/excel.slicer#style)|Constant value that represents the Slicer style. Possible values are: "SlicerStyleLight1" through "SlicerStyleLight6", "TableStyleOther1" through "TableStyleOther2", "SlicerStyleDark1" through "SlicerStyleDark6". A custom user-defined style present in the workbook can also be specified.|
||[top](/javascript/api/excel/excel.slicer#top)|Represents the distance, in points, from the top edge of the slicer to the top of the worksheet.|
||[width](/javascript/api/excel/excel.slicer#width)|Represents the width, in points, of the slicer.|
|[SlicerCollection](/javascript/api/excel/excel.slicercollection)|[add(slicerSource: string \| PivotTable \| Table, sourceField: string \| PivotField \| number \| TableColumn, slicerDestination?: string \| Worksheet)](/javascript/api/excel/excel.slicercollection#add-slicersource--sourcefield--slicerdestination-)|Adds a new slicer to the workbook.|
||[getCount()](/javascript/api/excel/excel.slicercollection#getcount--)|Returns the number of slicers in the collection.|
||[getItem(key: string)](/javascript/api/excel/excel.slicercollection#getitem-key-)|Gets a slicer object using its name or id.|
||[getItemAt(index: number)](/javascript/api/excel/excel.slicercollection#getitemat-index-)|Gets a slicer based on its position in the collection.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.slicercollection#getitemornullobject-key-)|Gets a slicer using its name or id. If the slicer does not exist, will return a null object.|
||[items](/javascript/api/excel/excel.slicercollection#items)|Gets the loaded child items in this collection.|
|[SlicerCollectionData](/javascript/api/excel/excel.slicercollectiondata)|[items](/javascript/api/excel/excel.slicercollectiondata#items)||
|[SlicerCollectionLoadOptions](/javascript/api/excel/excel.slicercollectionloadoptions)|[$all](/javascript/api/excel/excel.slicercollectionloadoptions#$all)||
||[caption](/javascript/api/excel/excel.slicercollectionloadoptions#caption)|For EACH ITEM in the collection: Represents the caption of slicer.|
||[height](/javascript/api/excel/excel.slicercollectionloadoptions#height)|For EACH ITEM in the collection: Represents the height, in points, of the slicer.|
||[id](/javascript/api/excel/excel.slicercollectionloadoptions#id)|For EACH ITEM in the collection: Represents the unique id of slicer. Read-only.|
||[isFilterCleared](/javascript/api/excel/excel.slicercollectionloadoptions#isfiltercleared)|For EACH ITEM in the collection: True if all filters currently applied on the slicer are cleared.|
||[left](/javascript/api/excel/excel.slicercollectionloadoptions#left)|For EACH ITEM in the collection: Represents the distance, in points, from the left side of the slicer to the left of the worksheet.|
||[name](/javascript/api/excel/excel.slicercollectionloadoptions#name)|For EACH ITEM in the collection: Represents the name of slicer.|
||[nameInFormula](/javascript/api/excel/excel.slicercollectionloadoptions#nameinformula)|For EACH ITEM in the collection: Represents the slicer name used in the formula.|
||[sortBy](/javascript/api/excel/excel.slicercollectionloadoptions#sortby)|For EACH ITEM in the collection: Represents the sort order of the items in the slicer. Possible values are: DataSourceOrder, Ascending, Descending.|
||[style](/javascript/api/excel/excel.slicercollectionloadoptions#style)|For EACH ITEM in the collection: Constant value that represents the Slicer style. Possible values are: "SlicerStyleLight1" through "SlicerStyleLight6", "TableStyleOther1" through "TableStyleOther2", "SlicerStyleDark1" through "SlicerStyleDark6". A custom user-defined style present in the workbook can also be specified.|
||[top](/javascript/api/excel/excel.slicercollectionloadoptions#top)|For EACH ITEM in the collection: Represents the distance, in points, from the top edge of the slicer to the top of the worksheet.|
||[width](/javascript/api/excel/excel.slicercollectionloadoptions#width)|For EACH ITEM in the collection: Represents the width, in points, of the slicer.|
||[worksheet](/javascript/api/excel/excel.slicercollectionloadoptions#worksheet)|For EACH ITEM in the collection: Represents the worksheet containing the slicer.|
|[SlicerCollectionUpdateData](/javascript/api/excel/excel.slicercollectionupdatedata)|[items](/javascript/api/excel/excel.slicercollectionupdatedata#items)||
|[SlicerData](/javascript/api/excel/excel.slicerdata)|[caption](/javascript/api/excel/excel.slicerdata#caption)|Represents the caption of slicer.|
||[height](/javascript/api/excel/excel.slicerdata#height)|Represents the height, in points, of the slicer.|
||[id](/javascript/api/excel/excel.slicerdata#id)|Represents the unique id of slicer. Read-only.|
||[isFilterCleared](/javascript/api/excel/excel.slicerdata#isfiltercleared)|True if all filters currently applied on the slicer are cleared.|
||[left](/javascript/api/excel/excel.slicerdata#left)|Represents the distance, in points, from the left side of the slicer to the left of the worksheet.|
||[name](/javascript/api/excel/excel.slicerdata#name)|Represents the name of slicer.|
||[nameInFormula](/javascript/api/excel/excel.slicerdata#nameinformula)|Represents the slicer name used in the formula.|
||[slicerItems](/javascript/api/excel/excel.slicerdata#sliceritems)|Represents the collection of SlicerItems that are part of the slicer. Read-only.|
||[sortBy](/javascript/api/excel/excel.slicerdata#sortby)|Represents the sort order of the items in the slicer. Possible values are: DataSourceOrder, Ascending, Descending.|
||[style](/javascript/api/excel/excel.slicerdata#style)|Constant value that represents the Slicer style. Possible values are: "SlicerStyleLight1" through "SlicerStyleLight6", "TableStyleOther1" through "TableStyleOther2", "SlicerStyleDark1" through "SlicerStyleDark6". A custom user-defined style present in the workbook can also be specified.|
||[top](/javascript/api/excel/excel.slicerdata#top)|Represents the distance, in points, from the top edge of the slicer to the top of the worksheet.|
||[width](/javascript/api/excel/excel.slicerdata#width)|Represents the width, in points, of the slicer.|
||[worksheet](/javascript/api/excel/excel.slicerdata#worksheet)|Represents the worksheet containing the slicer. Read-only.|
|[SlicerItem](/javascript/api/excel/excel.sliceritem)|[isSelected](/javascript/api/excel/excel.sliceritem#isselected)|True if the slicer item is selected.|
||[hasData](/javascript/api/excel/excel.sliceritem#hasdata)|True if the slicer item has data.|
||[key](/javascript/api/excel/excel.sliceritem#key)|Represents the unique value representing the slicer item.|
||[name](/javascript/api/excel/excel.sliceritem#name)|Represents the title displayed in the UI.|
||[set(properties: Excel.SlicerItem)](/javascript/api/excel/excel.sliceritem#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.SlicerItemUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.sliceritem#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
|[SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection)|[getCount()](/javascript/api/excel/excel.sliceritemcollection#getcount--)|Returns the number of slicer items in the slicer.|
||[getItem(key: string)](/javascript/api/excel/excel.sliceritemcollection#getitem-key-)|Gets a slicer item object using its key or name.|
||[getItemAt(index: number)](/javascript/api/excel/excel.sliceritemcollection#getitemat-index-)|Gets a slicer item based on its position in the collection.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.sliceritemcollection#getitemornullobject-key-)|Gets a slicer item using its key or name. If the slicer item does not exist, will return a null object.|
||[items](/javascript/api/excel/excel.sliceritemcollection#items)|Gets the loaded child items in this collection.|
|[SlicerItemCollectionData](/javascript/api/excel/excel.sliceritemcollectiondata)|[items](/javascript/api/excel/excel.sliceritemcollectiondata#items)||
|[SlicerItemCollectionLoadOptions](/javascript/api/excel/excel.sliceritemcollectionloadoptions)|[$all](/javascript/api/excel/excel.sliceritemcollectionloadoptions#$all)||
||[hasData](/javascript/api/excel/excel.sliceritemcollectionloadoptions#hasdata)|For EACH ITEM in the collection: True if the slicer item has data.|
||[isSelected](/javascript/api/excel/excel.sliceritemcollectionloadoptions#isselected)|For EACH ITEM in the collection: True if the slicer item is selected.|
||[key](/javascript/api/excel/excel.sliceritemcollectionloadoptions#key)|For EACH ITEM in the collection: Represents the unique value representing the slicer item.|
||[name](/javascript/api/excel/excel.sliceritemcollectionloadoptions#name)|For EACH ITEM in the collection: Represents the title displayed in the UI.|
|[SlicerItemCollectionUpdateData](/javascript/api/excel/excel.sliceritemcollectionupdatedata)|[items](/javascript/api/excel/excel.sliceritemcollectionupdatedata#items)||
|[SlicerItemData](/javascript/api/excel/excel.sliceritemdata)|[hasData](/javascript/api/excel/excel.sliceritemdata#hasdata)|True if the slicer item has data.|
||[isSelected](/javascript/api/excel/excel.sliceritemdata#isselected)|True if the slicer item is selected.|
||[key](/javascript/api/excel/excel.sliceritemdata#key)|Represents the unique value representing the slicer item.|
||[name](/javascript/api/excel/excel.sliceritemdata#name)|Represents the title displayed in the UI.|
|[SlicerItemLoadOptions](/javascript/api/excel/excel.sliceritemloadoptions)|[$all](/javascript/api/excel/excel.sliceritemloadoptions#$all)||
||[hasData](/javascript/api/excel/excel.sliceritemloadoptions#hasdata)|True if the slicer item has data.|
||[isSelected](/javascript/api/excel/excel.sliceritemloadoptions#isselected)|True if the slicer item is selected.|
||[key](/javascript/api/excel/excel.sliceritemloadoptions#key)|Represents the unique value representing the slicer item.|
||[name](/javascript/api/excel/excel.sliceritemloadoptions#name)|Represents the title displayed in the UI.|
|[SlicerItemUpdateData](/javascript/api/excel/excel.sliceritemupdatedata)|[isSelected](/javascript/api/excel/excel.sliceritemupdatedata#isselected)|True if the slicer item is selected.|
|[SlicerLoadOptions](/javascript/api/excel/excel.slicerloadoptions)|[$all](/javascript/api/excel/excel.slicerloadoptions#$all)||
||[caption](/javascript/api/excel/excel.slicerloadoptions#caption)|Represents the caption of slicer.|
||[height](/javascript/api/excel/excel.slicerloadoptions#height)|Represents the height, in points, of the slicer.|
||[id](/javascript/api/excel/excel.slicerloadoptions#id)|Represents the unique id of slicer. Read-only.|
||[isFilterCleared](/javascript/api/excel/excel.slicerloadoptions#isfiltercleared)|True if all filters currently applied on the slicer are cleared.|
||[left](/javascript/api/excel/excel.slicerloadoptions#left)|Represents the distance, in points, from the left side of the slicer to the left of the worksheet.|
||[name](/javascript/api/excel/excel.slicerloadoptions#name)|Represents the name of slicer.|
||[nameInFormula](/javascript/api/excel/excel.slicerloadoptions#nameinformula)|Represents the slicer name used in the formula.|
||[sortBy](/javascript/api/excel/excel.slicerloadoptions#sortby)|Represents the sort order of the items in the slicer. Possible values are: DataSourceOrder, Ascending, Descending.|
||[style](/javascript/api/excel/excel.slicerloadoptions#style)|Constant value that represents the Slicer style. Possible values are: "SlicerStyleLight1" through "SlicerStyleLight6", "TableStyleOther1" through "TableStyleOther2", "SlicerStyleDark1" through "SlicerStyleDark6". A custom user-defined style present in the workbook can also be specified.|
||[top](/javascript/api/excel/excel.slicerloadoptions#top)|Represents the distance, in points, from the top edge of the slicer to the top of the worksheet.|
||[width](/javascript/api/excel/excel.slicerloadoptions#width)|Represents the width, in points, of the slicer.|
||[worksheet](/javascript/api/excel/excel.slicerloadoptions#worksheet)|Represents the worksheet containing the slicer.|
|[SlicerStyle](/javascript/api/excel/excel.slicerstyle)|[delete()](/javascript/api/excel/excel.slicerstyle#delete--)|Deletes the SlicerStyle.|
||[duplicate()](/javascript/api/excel/excel.slicerstyle#duplicate--)|Creates a duplicate of this SlicerStyle with copies of all the style elements.|
||[name](/javascript/api/excel/excel.slicerstyle#name)|Gets the name of the SlicerStyle.|
||[readOnly](/javascript/api/excel/excel.slicerstyle#readonly)|Specifies if this SlicerStyle object is read-only. Read-only.|
||[set(properties: Excel.SlicerStyle)](/javascript/api/excel/excel.slicerstyle#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.SlicerStyleUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.slicerstyle#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
|[SlicerStyleCollection](/javascript/api/excel/excel.slicerstylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.slicerstylecollection#add-name--makeuniquename-)|Creates a blank SlicerStyle with the specified name.|
||[getCount()](/javascript/api/excel/excel.slicerstylecollection#getcount--)|Gets the number of slicer styles in the collection.|
||[getDefault()](/javascript/api/excel/excel.slicerstylecollection#getdefault--)|Gets the default SlicerStyle for the parent object's scope.|
||[getItem(name: string)](/javascript/api/excel/excel.slicerstylecollection#getitem-name-)|Gets a SlicerStyle by name.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.slicerstylecollection#getitemornullobject-name-)|Gets a SlicerStyle by name. If the SlicerStyle does not exist, will return a null object.|
||[items](/javascript/api/excel/excel.slicerstylecollection#items)|Gets the loaded child items in this collection.|
||[setDefault(newDefaultStyle: SlicerStyle \| string)](/javascript/api/excel/excel.slicerstylecollection#setdefault-newdefaultstyle-)|Sets the default SlicerStyle for use in the parent object's scope.|
|[SlicerStyleCollectionData](/javascript/api/excel/excel.slicerstylecollectiondata)|[items](/javascript/api/excel/excel.slicerstylecollectiondata#items)||
|[SlicerStyleCollectionLoadOptions](/javascript/api/excel/excel.slicerstylecollectionloadoptions)|[$all](/javascript/api/excel/excel.slicerstylecollectionloadoptions#$all)||
||[name](/javascript/api/excel/excel.slicerstylecollectionloadoptions#name)|For EACH ITEM in the collection: Gets the name of the SlicerStyle.|
||[readOnly](/javascript/api/excel/excel.slicerstylecollectionloadoptions#readonly)|For EACH ITEM in the collection: Specifies if this SlicerStyle object is read-only. Read-only.|
|[SlicerStyleCollectionUpdateData](/javascript/api/excel/excel.slicerstylecollectionupdatedata)|[items](/javascript/api/excel/excel.slicerstylecollectionupdatedata#items)||
|[SlicerStyleData](/javascript/api/excel/excel.slicerstyledata)|[name](/javascript/api/excel/excel.slicerstyledata#name)|Gets the name of the SlicerStyle.|
||[readOnly](/javascript/api/excel/excel.slicerstyledata#readonly)|Specifies if this SlicerStyle object is read-only. Read-only.|
|[SlicerStyleLoadOptions](/javascript/api/excel/excel.slicerstyleloadoptions)|[$all](/javascript/api/excel/excel.slicerstyleloadoptions#$all)||
||[name](/javascript/api/excel/excel.slicerstyleloadoptions#name)|Gets the name of the SlicerStyle.|
||[readOnly](/javascript/api/excel/excel.slicerstyleloadoptions#readonly)|Specifies if this SlicerStyle object is read-only. Read-only.|
|[SlicerStyleUpdateData](/javascript/api/excel/excel.slicerstyleupdatedata)|[name](/javascript/api/excel/excel.slicerstyleupdatedata#name)|Gets the name of the SlicerStyle.|
|[SlicerUpdateData](/javascript/api/excel/excel.slicerupdatedata)|[caption](/javascript/api/excel/excel.slicerupdatedata#caption)|Represents the caption of slicer.|
||[height](/javascript/api/excel/excel.slicerupdatedata#height)|Represents the height, in points, of the slicer.|
||[left](/javascript/api/excel/excel.slicerupdatedata#left)|Represents the distance, in points, from the left side of the slicer to the left of the worksheet.|
||[name](/javascript/api/excel/excel.slicerupdatedata#name)|Represents the name of slicer.|
||[nameInFormula](/javascript/api/excel/excel.slicerupdatedata#nameinformula)|Represents the slicer name used in the formula.|
||[sortBy](/javascript/api/excel/excel.slicerupdatedata#sortby)|Represents the sort order of the items in the slicer. Possible values are: DataSourceOrder, Ascending, Descending.|
||[style](/javascript/api/excel/excel.slicerupdatedata#style)|Constant value that represents the Slicer style. Possible values are: "SlicerStyleLight1" through "SlicerStyleLight6", "TableStyleOther1" through "TableStyleOther2", "SlicerStyleDark1" through "SlicerStyleDark6". A custom user-defined style present in the workbook can also be specified.|
||[top](/javascript/api/excel/excel.slicerupdatedata#top)|Represents the distance, in points, from the top edge of the slicer to the top of the worksheet.|
||[width](/javascript/api/excel/excel.slicerupdatedata#width)|Represents the width, in points, of the slicer.|
||[worksheet](/javascript/api/excel/excel.slicerupdatedata#worksheet)|Represents the worksheet containing the slicer.|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|Changes the table to use the default table style.|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|Occurs when filter is applied on a specific table.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|Occurs when filter is applied on any table in a workbook, or a worksheet.|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|Represents the id of the table in which the filter is applied..|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|Represents the type of the event. See Excel.EventType for details.|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|Represents the id of the worksheet which contains the table.|
|[TableStyle](/javascript/api/excel/excel.tablestyle)|[delete()](/javascript/api/excel/excel.tablestyle#delete--)|Deletes the TableStyle.|
||[duplicate()](/javascript/api/excel/excel.tablestyle#duplicate--)|Creates a duplicate of this TableStyle with copies of all the style elements.|
||[name](/javascript/api/excel/excel.tablestyle#name)|Gets the name of the TableStyle.|
||[readOnly](/javascript/api/excel/excel.tablestyle#readonly)|Specifies if this TableStyle object is read-only. Read-only.|
||[set(properties: Excel.TableStyle)](/javascript/api/excel/excel.tablestyle#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.TableStyleUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.tablestyle#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
|[TableStyleCollection](/javascript/api/excel/excel.tablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.tablestylecollection#add-name--makeuniquename-)|Creates a blank TableStyle with the specified name.|
||[getCount()](/javascript/api/excel/excel.tablestylecollection#getcount--)|Gets the number of table styles in the collection.|
||[getDefault()](/javascript/api/excel/excel.tablestylecollection#getdefault--)|Gets the default TableStyle for the parent object's scope.|
||[getItem(name: string)](/javascript/api/excel/excel.tablestylecollection#getitem-name-)|Gets a TableStyle by name.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.tablestylecollection#getitemornullobject-name-)|Gets a TableStyle by name. If the TableStyle does not exist, will return a null object.|
||[items](/javascript/api/excel/excel.tablestylecollection#items)|Gets the loaded child items in this collection.|
||[setDefault(newDefaultStyle: TableStyle \| string)](/javascript/api/excel/excel.tablestylecollection#setdefault-newdefaultstyle-)|Sets the default TableStyle for use in the parent object's scope.|
|[TableStyleCollectionData](/javascript/api/excel/excel.tablestylecollectiondata)|[items](/javascript/api/excel/excel.tablestylecollectiondata#items)||
|[TableStyleCollectionLoadOptions](/javascript/api/excel/excel.tablestylecollectionloadoptions)|[$all](/javascript/api/excel/excel.tablestylecollectionloadoptions#$all)||
||[name](/javascript/api/excel/excel.tablestylecollectionloadoptions#name)|For EACH ITEM in the collection: Gets the name of the TableStyle.|
||[readOnly](/javascript/api/excel/excel.tablestylecollectionloadoptions#readonly)|For EACH ITEM in the collection: Specifies if this TableStyle object is read-only. Read-only.|
|[TableStyleCollectionUpdateData](/javascript/api/excel/excel.tablestylecollectionupdatedata)|[items](/javascript/api/excel/excel.tablestylecollectionupdatedata#items)||
|[TableStyleData](/javascript/api/excel/excel.tablestyledata)|[name](/javascript/api/excel/excel.tablestyledata#name)|Gets the name of the TableStyle.|
||[readOnly](/javascript/api/excel/excel.tablestyledata#readonly)|Specifies if this TableStyle object is read-only. Read-only.|
|[TableStyleLoadOptions](/javascript/api/excel/excel.tablestyleloadoptions)|[$all](/javascript/api/excel/excel.tablestyleloadoptions#$all)||
||[name](/javascript/api/excel/excel.tablestyleloadoptions#name)|Gets the name of the TableStyle.|
||[readOnly](/javascript/api/excel/excel.tablestyleloadoptions#readonly)|Specifies if this TableStyle object is read-only. Read-only.|
|[TableStyleUpdateData](/javascript/api/excel/excel.tablestyleupdatedata)|[name](/javascript/api/excel/excel.tablestyleupdatedata#name)|Gets the name of the TableStyle.|
|[TimelineStyle](/javascript/api/excel/excel.timelinestyle)|[delete()](/javascript/api/excel/excel.timelinestyle#delete--)|Deletes the TableStyle.|
||[duplicate()](/javascript/api/excel/excel.timelinestyle#duplicate--)|Creates a duplicate of this TimelineStyle with copies of all the style elements.|
||[name](/javascript/api/excel/excel.timelinestyle#name)|Gets the name of the TimelineStyle.|
||[readOnly](/javascript/api/excel/excel.timelinestyle#readonly)|Specifies if this TimelineStyle object is read-only. Read-only.|
||[set(properties: Excel.TimelineStyle)](/javascript/api/excel/excel.timelinestyle#set-properties-)|Sets multiple properties on the object at the same time, based on an existing loaded object.|
||[set(properties: Interfaces.TimelineStyleUpdateData, options?: OfficeExtension.UpdateOptions)](/javascript/api/excel/excel.timelinestyle#set-properties--options-)|Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.|
|[TimelineStyleCollection](/javascript/api/excel/excel.timelinestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.timelinestylecollection#add-name--makeuniquename-)|Creates a blank TimelineStyle with the specified name.|
||[getCount()](/javascript/api/excel/excel.timelinestylecollection#getcount--)|Gets the number of timeline styles in the collection.|
||[getDefault()](/javascript/api/excel/excel.timelinestylecollection#getdefault--)|Gets the default TimelineStyle for the parent object's scope.|
||[getItem(name: string)](/javascript/api/excel/excel.timelinestylecollection#getitem-name-)|Gets a TimelineStyle by name.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.timelinestylecollection#getitemornullobject-name-)|Gets a TimelineStyle by name. If the TimelineStyle does not exist, will return a null object.|
||[items](/javascript/api/excel/excel.timelinestylecollection#items)|Gets the loaded child items in this collection.|
||[setDefault(newDefaultStyle: TimelineStyle \| string)](/javascript/api/excel/excel.timelinestylecollection#setdefault-newdefaultstyle-)|Sets the default TimelineStyle for use in the parent object's scope.|
|[TimelineStyleCollectionData](/javascript/api/excel/excel.timelinestylecollectiondata)|[items](/javascript/api/excel/excel.timelinestylecollectiondata#items)||
|[TimelineStyleCollectionLoadOptions](/javascript/api/excel/excel.timelinestylecollectionloadoptions)|[$all](/javascript/api/excel/excel.timelinestylecollectionloadoptions#$all)||
||[name](/javascript/api/excel/excel.timelinestylecollectionloadoptions#name)|For EACH ITEM in the collection: Gets the name of the TimelineStyle.|
||[readOnly](/javascript/api/excel/excel.timelinestylecollectionloadoptions#readonly)|For EACH ITEM in the collection: Specifies if this TimelineStyle object is read-only. Read-only.|
|[TimelineStyleCollectionUpdateData](/javascript/api/excel/excel.timelinestylecollectionupdatedata)|[items](/javascript/api/excel/excel.timelinestylecollectionupdatedata#items)||
|[TimelineStyleData](/javascript/api/excel/excel.timelinestyledata)|[name](/javascript/api/excel/excel.timelinestyledata#name)|Gets the name of the TimelineStyle.|
||[readOnly](/javascript/api/excel/excel.timelinestyledata#readonly)|Specifies if this TimelineStyle object is read-only. Read-only.|
|[TimelineStyleLoadOptions](/javascript/api/excel/excel.timelinestyleloadoptions)|[$all](/javascript/api/excel/excel.timelinestyleloadoptions#$all)||
||[name](/javascript/api/excel/excel.timelinestyleloadoptions#name)|Gets the name of the TimelineStyle.|
||[readOnly](/javascript/api/excel/excel.timelinestyleloadoptions#readonly)|Specifies if this TimelineStyle object is read-only. Read-only.|
|[TimelineStyleUpdateData](/javascript/api/excel/excel.timelinestyleupdatedata)|[name](/javascript/api/excel/excel.timelinestyleupdatedata#name)|Gets the name of the TimelineStyle.|
|[Workbook](/javascript/api/excel/excel.workbook)|[close(closeBehavior?: "Save" \| "SkipSave")](/javascript/api/excel/excel.workbook#close-closebehavior-)|Close current workbook.|
||[close(closeBehavior?: Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#close-closebehavior-)|Close current workbook.|
||[getActiveSlicer()](/javascript/api/excel/excel.workbook#getactiveslicer--)|Gets the currently active slicer in the workbook. If there is no active slicer, an `ItemNotFound` exception is thrown.|
||[getActiveSlicerOrNullObject()](/javascript/api/excel/excel.workbook#getactiveslicerornullobject--)|Gets the currently active slicer in the workbook. If there is no active slicer, a null object is returned.|
||[comments](/javascript/api/excel/excel.workbook#comments)|Represents a collection of Comments associated with the workbook. Read-only.|
||[pivotTableStyles](/javascript/api/excel/excel.workbook#pivottablestyles)|Represents a collection of PivotTableStyles associated with the workbook. Read-only.|
||[slicerStyles](/javascript/api/excel/excel.workbook#slicerstyles)|Represents a collection of SlicerStyles associated with the workbook. Read-only.|
||[slicers](/javascript/api/excel/excel.workbook#slicers)|Represents a collection of Slicers associated with the workbook. Read-only.|
||[tableStyles](/javascript/api/excel/excel.workbook#tablestyles)|Represents a collection of TableStyles associated with the workbook. Read-only.|
||[timelineStyles](/javascript/api/excel/excel.workbook#timelinestyles)|Represents a collection of TimelineStyles associated with the workbook. Read-only.|
||[save(saveBehavior?: "Save" \| "Prompt")](/javascript/api/excel/excel.workbook#save-savebehavior-)|Save current workbook.|
||[save(saveBehavior?: Excel.SaveBehavior)](/javascript/api/excel/excel.workbook#save-savebehavior-)|Save current workbook.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|True if the workbook uses the 1904 date system.|
|[WorkbookData](/javascript/api/excel/excel.workbookdata)|[comments](/javascript/api/excel/excel.workbookdata#comments)|Represents a collection of Comments associated with the workbook. Read-only.|
||[pivotTableStyles](/javascript/api/excel/excel.workbookdata#pivottablestyles)|Represents a collection of PivotTableStyles associated with the workbook. Read-only.|
||[slicerStyles](/javascript/api/excel/excel.workbookdata#slicerstyles)|Represents a collection of SlicerStyles associated with the workbook. Read-only.|
||[slicers](/javascript/api/excel/excel.workbookdata#slicers)|Represents a collection of Slicers associated with the workbook. Read-only.|
||[tableStyles](/javascript/api/excel/excel.workbookdata#tablestyles)|Represents a collection of TableStyles associated with the workbook. Read-only.|
||[timelineStyles](/javascript/api/excel/excel.workbookdata#timelinestyles)|Represents a collection of TimelineStyles associated with the workbook. Read-only.|
||[use1904DateSystem](/javascript/api/excel/excel.workbookdata#use1904datesystem)|True if the workbook uses the 1904 date system.|
|[WorkbookLoadOptions](/javascript/api/excel/excel.workbookloadoptions)|[use1904DateSystem](/javascript/api/excel/excel.workbookloadoptions#use1904datesystem)|True if the workbook uses the 1904 date system.|
|[WorkbookUpdateData](/javascript/api/excel/excel.workbookupdatedata)|[use1904DateSystem](/javascript/api/excel/excel.workbookupdatedata#use1904datesystem)|True if the workbook uses the 1904 date system.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[comments](/javascript/api/excel/excel.worksheet#comments)|Returns a collection of all the Comments objects on the worksheet. Read-only.|
||[onColumnSorted](/javascript/api/excel/excel.worksheet#oncolumnsorted)|Occurs when sorting on columns.|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|Occurs when filter is applied on a specific worksheet.|
||[onRowHiddenChanged](/javascript/api/excel/excel.worksheet#onrowhiddenchanged)|Occurs when row hidden state changed on a specific worksheet.|
||[onRowSorted](/javascript/api/excel/excel.worksheet#onrowsorted)|Occurs when sorting on rows.|
||[onSingleClicked](/javascript/api/excel/excel.worksheet#onsingleclicked)|Occurs when left-clicked/tapped operation happens in the worksheet.|
||[slicers](/javascript/api/excel/excel.worksheet#slicers)|Returns collection of slicers that are part of the worksheet. Read-only.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: "None" \| "Before" \| "After" \| "Beginning" \| "End", relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|Inserts the specified worksheets of a workbook into the current workbook.|
||[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|Inserts the specified worksheets of a workbook into the current workbook.|
||[onColumnSorted](/javascript/api/excel/excel.worksheetcollection#oncolumnsorted)|Occurs when sorting on columns.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|Occurs when any worksheet's filter is applied in the workbook.|
||[onRowHiddenChanged](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged)|Occurs when any worksheet in the workbook has row hidden state changed.|
||[onRowSorted](/javascript/api/excel/excel.worksheetcollection#onrowsorted)|Occurs when sorting on rows.|
||[onSingleClicked](/javascript/api/excel/excel.worksheetcollection#onsingleclicked)|Occurs when left-clicked/tapped operation happens in the worksheet collection.|
|[WorksheetColumnSortedEventArgs](/javascript/api/excel/excel.worksheetcolumnsortedeventargs)|[address](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#address)|Gets the range address that represents the sorted areas of a specific worksheet.|
||[source](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#source)|Gets the source of the event. See Excel.EventSource for details.|
||[type](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#type)|Gets the type of the event. See Excel.EventType for details.|
||[worksheetId](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#worksheetid)|Gets the id of the worksheet where the sorting happened.|
|[WorksheetData](/javascript/api/excel/excel.worksheetdata)|[comments](/javascript/api/excel/excel.worksheetdata#comments)|Returns a collection of all the Comments objects on the worksheet. Read-only.|
||[slicers](/javascript/api/excel/excel.worksheetdata#slicers)|Returns collection of slicers that are part of the worksheet. Read-only.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|Represents the type of the event. See Excel.EventType for details.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|Represents the id of the worksheet in which the filter is applied.|
|[WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs)|[address](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#address)|Gets the range address that represents the changed area of a specific worksheet.|
||[changeType](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#changetype)|Gets the change type that represents how the Changed event is triggered. See Excel.RowHiddenChangeType for details.|
||[source](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#source)|Gets the source of the event. See Excel.EventSource for details.|
||[type](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#type)|Gets the type of the event. See Excel.EventType for details.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#worksheetid)|Gets the id of the worksheet in which the data changed.|
|[WorksheetRowSortedEventArgs](/javascript/api/excel/excel.worksheetrowsortedeventargs)|[address](/javascript/api/excel/excel.worksheetrowsortedeventargs#address)|Gets the range address that represents the sorted areas of a specific worksheet.|
||[source](/javascript/api/excel/excel.worksheetrowsortedeventargs#source)|Gets the source of the event. See Excel.EventSource for details.|
||[type](/javascript/api/excel/excel.worksheetrowsortedeventargs#type)|Gets the type of the event. See Excel.EventType for details.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowsortedeventargs#worksheetid)|Gets the id of the worksheet where the sorting happened.|
|[WorksheetSingleClickedEventArgs](/javascript/api/excel/excel.worksheetsingleclickedeventargs)|[address](/javascript/api/excel/excel.worksheetsingleclickedeventargs#address)|Gets the address that represents the cell which was left-clicked/tapped for a specific worksheet.|
||[offsetX](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsetx)|The distance, in points, from the left-clicked/tapped point to the left (right for RTL) gridline edge of the left-clicked/tapped cell.|
||[offsetY](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsety)|The distance, in points, from the left-clicked/tapped point to the top gridline edge of the left-clicked/tapped cell.|
||[type](/javascript/api/excel/excel.worksheetsingleclickedeventargs#type)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetsingleclickedeventargs#worksheetid)|Gets the id of the worksheet in which the cell was left-clicked/tapped.|

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel)
- [Excel JavaScript API requirement sets](./excel-api-requirement-sets.md)
