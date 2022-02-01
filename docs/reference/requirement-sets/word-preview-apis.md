---
title: Word JavaScript preview APIs
description: Details about upcoming Word JavaScript APIs.
ms.date: 12/14/2021
ms.prod: word
ms.localizationpriority: medium
---

# Word JavaScript preview APIs

New Word JavaScript APIs are first introduced in "preview" and later become part of a specific, numbered requirement set after sufficient testing occurs and user feedback is acquired.

[!INCLUDE [Information about using Word preview APIs](../../includes/word-preview-apis-note.md)]
[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

## API list

The following table lists the Word JavaScript APIs currently in preview, except those that are [available only in Word on the web](#web-only-api-list). To see a complete list of all Word JavaScript APIs (including preview APIs and previously released APIs), see [all Word JavaScript APIs](/javascript/api/word?view=word-js-preview&preserve-view=true).

| Class | Fields | Description |
|:---|:---|:---|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[onDataChanged](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-onDataChanged-member)|Occurs when data within the content control are changed.|
||[onDeleted](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-onDeleted-member)|Occurs when the content control is deleted.|
||[onSelectionChanged](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-onSelectionChanged-member)|Occurs when selection within the content control is changed.|
|[ContentControlEventArgs](/javascript/api/word/word.contentcontroleventargs)|[contentControl](/javascript/api/word/word.contentcontroleventargs#word-word-contentcontroleventargs-contentControl-member)|The object that raised the event.|
||[eventType](/javascript/api/word/word.contentcontroleventargs#word-word-contentcontroleventargs-eventType-member)|The event type.|
|[CustomXmlPart](/javascript/api/word/word.customxmlpart)|[delete()](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-delete-member(1))|Deletes the custom XML part.|
||[deleteAttribute(xpath: string, namespaceMappings: any, name: string)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-deleteAttribute-member(1))|Deletes an attribute with the given name from the element identified by xpath.|
||[deleteElement(xpath: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-deleteElement-member(1))|Deletes the element identified by xpath.|
||[getXml()](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-getXml-member(1))|Gets the full XML content of the custom XML part.|
||[id](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-id-member)|Gets the ID of the custom XML part.|
||[insertAttribute(xpath: string, namespaceMappings: any, name: string, value: string)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-insertAttribute-member(1))|Inserts an attribute with the given name and value to the element identified by xpath.|
||[insertElement(xpath: string, xml: string, namespaceMappings: any, index?: number)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-insertElement-member(1))|Inserts the given XML under the parent element identified by xpath at child position index.|
||[namespaceUri](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-namespaceUri-member)|Gets the namespace URI of the custom XML part.|
||[query(xpath: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-query-member(1))|Queries the XML content of the custom XML part.|
||[setXml(xml: string)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-setXml-member(1))|Sets the full XML content of the custom XML part.|
||[updateAttribute(xpath: string, namespaceMappings: any, name: string, value: string)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-updateAttribute-member(1))|Updates the value of an attribute with the given name of the element identified by xpath.|
||[updateElement(xpath: string, xml: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-updateElement-member(1))|Updates the XML of the element identified by xpath.|
|[CustomXmlPartCollection](/javascript/api/word/word.customxmlpartcollection)|[add(xml: string)](/javascript/api/word/word.customxmlpartcollection#word-word-customxmlpartcollection-add-member(1))|Adds a new custom XML part to the document.|
||[getByNamespace(namespaceUri: string)](/javascript/api/word/word.customxmlpartcollection#word-word-customxmlpartcollection-getByNamespace-member(1))|Gets a new scoped collection of custom XML parts whose namespaces match the given namespace.|
||[getCount()](/javascript/api/word/word.customxmlpartcollection#word-word-customxmlpartcollection-getCount-member(1))|Gets the number of items in the collection.|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartcollection#word-word-customxmlpartcollection-getItem-member(1))|Gets a custom XML part based on its ID.|
||[getItemOrNullObject(id: string)](/javascript/api/word/word.customxmlpartcollection#word-word-customxmlpartcollection-getItemOrNullObject-member(1))|Gets a custom XML part based on its ID.|
||[items](/javascript/api/word/word.customxmlpartcollection#word-word-customxmlpartcollection-items-member)|Gets the loaded child items in this collection.|
|[CustomXmlPartScopedCollection](/javascript/api/word/word.customxmlpartscopedcollection)|[getCount()](/javascript/api/word/word.customxmlpartscopedcollection#word-word-customxmlpartscopedcollection-getCount-member(1))|Gets the number of items in the collection.|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartscopedcollection#word-word-customxmlpartscopedcollection-getItem-member(1))|Gets a custom XML part based on its ID.|
||[getItemOrNullObject(id: string)](/javascript/api/word/word.customxmlpartscopedcollection#word-word-customxmlpartscopedcollection-getItemOrNullObject-member(1))|Gets a custom XML part based on its ID.|
||[getOnlyItem()](/javascript/api/word/word.customxmlpartscopedcollection#word-word-customxmlpartscopedcollection-getOnlyItem-member(1))|If the collection contains exactly one item, this method returns it.|
||[getOnlyItemOrNullObject()](/javascript/api/word/word.customxmlpartscopedcollection#word-word-customxmlpartscopedcollection-getOnlyItemOrNullObject-member(1))|If the collection contains exactly one item, this method returns it.|
||[items](/javascript/api/word/word.customxmlpartscopedcollection#word-word-customxmlpartscopedcollection-items-member)|Gets the loaded child items in this collection.|
|[Document](/javascript/api/word/word.document)|[customXmlParts](/javascript/api/word/word.document#word-word-document-customXmlParts-member)|Gets the custom XML parts in the document.|
||[deleteBookmark(name: string)](/javascript/api/word/word.document#word-word-document-deleteBookmark-member(1))|Deletes a bookmark, if it exists, from the document.|
||[getBookmarkRange(name: string)](/javascript/api/word/word.document#word-word-document-getBookmarkRange-member(1))|Gets a bookmark's range.|
||[getBookmarkRangeOrNullObject(name: string)](/javascript/api/word/word.document#word-word-document-getBookmarkRangeOrNullObject-member(1))|Gets a bookmark's range.|
||[ignorePunct](/javascript/api/word/word.document#word-word-document-ignorePunct-member)||
||[ignoreSpace](/javascript/api/word/word.document#word-word-document-ignoreSpace-member)||
||[matchCase](/javascript/api/word/word.document#word-word-document-matchCase-member)||
||[matchPrefix](/javascript/api/word/word.document#word-word-document-matchPrefix-member)||
||[matchSuffix](/javascript/api/word/word.document#word-word-document-matchSuffix-member)||
||[matchWholeWord](/javascript/api/word/word.document#word-word-document-matchWholeWord-member)||
||[matchWildcards](/javascript/api/word/word.document#word-word-document-matchWildcards-member)||
||[onContentControlAdded](/javascript/api/word/word.document#word-word-document-onContentControlAdded-member)|Occurs when a content control is added.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| {            ignorePunct?: boolean            ignoreSpace?: boolean            matchCase?: boolean            matchPrefix?: boolean            matchSuffix?: boolean            matchWholeWord?: boolean            matchWildcards?: boolean        })](/javascript/api/word/word.document#word-word-document-search-member(1))|Performs a search with the specified search options on the scope of the whole document.|
||[settings](/javascript/api/word/word.document#word-word-document-settings-member)|Gets the add-in's settings in the document.|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[customXmlParts](/javascript/api/word/word.documentcreated#word-word-documentcreated-customXmlParts-member)|Gets the custom XML parts in the document.|
||[deleteBookmark(name: string)](/javascript/api/word/word.documentcreated#word-word-documentcreated-deleteBookmark-member(1))|Deletes a bookmark, if it exists, from the document.|
||[getBookmarkRange(name: string)](/javascript/api/word/word.documentcreated#word-word-documentcreated-getBookmarkRange-member(1))|Gets a bookmark's range.|
||[getBookmarkRangeOrNullObject(name: string)](/javascript/api/word/word.documentcreated#word-word-documentcreated-getBookmarkRangeOrNullObject-member(1))|Gets a bookmark's range.|
||[settings](/javascript/api/word/word.documentcreated#word-word-documentcreated-settings-member)|Gets the add-in's settings in the document.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[imageFormat](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-imageFormat-member)|Gets the format of the inline image.|
|[List](/javascript/api/word/word.list)|[getLevelFont(level: number)](/javascript/api/word/word.list#word-word-list-getLevelFont-member(1))|Gets the font of the bullet, number, or picture at the specified level in the list.|
||[getLevelPicture(level: number)](/javascript/api/word/word.list#word-word-list-getLevelPicture-member(1))|Gets the base64 encoded string representation of the picture at the specified level in the list.|
||[resetLevelFont(level: number, resetFontName?: boolean)](/javascript/api/word/word.list#word-word-list-resetLevelFont-member(1))|Resets the font of the bullet, number, or picture at the specified level in the list.|
||[setLevelPicture(level: number, base64EncodedImage?: string)](/javascript/api/word/word.list#word-word-list-setLevelPicture-member(1))|Sets the picture at the specified level in the list.|
|[Range](/javascript/api/word/word.range)|[getBookmarks(includeHidden?: boolean, includeAdjacent?: boolean)](/javascript/api/word/word.range#word-word-range-getBookmarks-member(1))|Gets the names all bookmarks in or overlapping the range.|
||[insertBookmark(name: string)](/javascript/api/word/word.range#word-word-range-insertBookmark-member(1))|Inserts a bookmark on the range.|
|[Setting](/javascript/api/word/word.setting)|[delete()](/javascript/api/word/word.setting#word-word-setting-delete-member(1))|Deletes the setting.|
||[key](/javascript/api/word/word.setting#word-word-setting-key-member)|Gets the key of the setting.|
||[value](/javascript/api/word/word.setting#word-word-setting-value-member)|Gets or sets the value of the setting.|
|[SettingCollection](/javascript/api/word/word.settingcollection)|[add(key: string, value: any)](/javascript/api/word/word.settingcollection#word-word-settingcollection-add-member(1))|Creates a new setting or sets an existing setting.|
||[deleteAll()](/javascript/api/word/word.settingcollection#word-word-settingcollection-deleteAll-member(1))|Deletes all settings in this add-in.|
||[getCount()](/javascript/api/word/word.settingcollection#word-word-settingcollection-getCount-member(1))|Gets the count of settings.|
||[getItem(key: string)](/javascript/api/word/word.settingcollection#word-word-settingcollection-getItem-member(1))|Gets a setting object by its key, which is case-sensitive.|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.settingcollection#word-word-settingcollection-getItemOrNullObject-member(1))|Gets a setting object by its key, which is case-sensitive.|
||[items](/javascript/api/word/word.settingcollection#word-word-settingcollection-items-member)|Gets the loaded child items in this collection.|
|[Table](/javascript/api/word/word.table)|[mergeCells(topRow: number, firstCell: number, bottomRow: number, lastCell: number)](/javascript/api/word/word.table#word-word-table-mergeCells-member(1))|Merges the cells bounded inclusively by a first and last cell.|
|[TableCell](/javascript/api/word/word.tablecell)|[split(rowCount: number, columnCount: number)](/javascript/api/word/word.tablecell#word-word-tablecell-split-member(1))|Splits the cell into the specified number of rows and columns.|
|[TableRow](/javascript/api/word/word.tablerow)|[insertContentControl()](/javascript/api/word/word.tablerow#word-word-tablerow-insertContentControl-member(1))|Inserts a content control on the row.|
||[merge()](/javascript/api/word/word.tablerow#word-word-tablerow-merge-member(1))|Merges the row into one cell.|

## Web-only API list

The following table lists the Word JavaScript APIs currently in preview only in Word on the web. To see a complete list of all Word JavaScript APIs (including preview APIs and previously released APIs), see [all Word JavaScript APIs](/javascript/api/word?view=word-js-preview&preserve-view=true).

| Class | Fields | Description |
|:---|:---|:---|
|[Body](/javascript/api/word/word.body)|[endnotes](/javascript/api/word/word.body#word-word-body-endnotes-member)|Gets the collection of endnotes in the body.|
||[footnotes](/javascript/api/word/word.body#word-word-body-footnotes-member)|Gets the collection of footnotes in the body.|
||[getComments()](/javascript/api/word/word.body#word-word-body-getComments-member(1))|Gets comments associated with the body.|
||[getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion)](/javascript/api/word/word.body#word-word-body-getReviewedText-member(1))|Gets reviewed text based on ChangeTrackingVersion selection.|
||[type](/javascript/api/word/word.body#word-word-body-type-member)|Gets the type of the body.|
|[Comment](/javascript/api/word/word.comment)|[authorEmail](/javascript/api/word/word.comment#word-word-comment-authorEmail-member)|Gets the email of the comment's author.|
||[authorName](/javascript/api/word/word.comment#word-word-comment-authorName-member)|Gets the name of the comment's author.|
||[content](/javascript/api/word/word.comment#word-word-comment-content-member)|Gets or sets the comment's content as plain text.|
||[creationDate](/javascript/api/word/word.comment#word-word-comment-creationDate-member)|Gets the creation date of the comment.|
||[delete()](/javascript/api/word/word.comment#word-word-comment-delete-member(1))|Deletes the comment and its replies.|
||[getRange()](/javascript/api/word/word.comment#word-word-comment-getRange-member(1))|Gets the range in the main document where the comment is on.|
||[id](/javascript/api/word/word.comment#word-word-comment-id-member)|ID|
||[replies](/javascript/api/word/word.comment#word-word-comment-replies-member)|Gets the collection of reply objects associated with the comment.|
||[reply(replyText: string)](/javascript/api/word/word.comment#word-word-comment-reply-member(1))|Adds a new reply to the end of the comment thread.|
||[resolved](/javascript/api/word/word.comment#word-word-comment-resolved-member)|Gets or sets the comment thread status.|
|[CommentCollection](/javascript/api/word/word.commentcollection)|[getFirst()](/javascript/api/word/word.commentcollection#word-word-commentcollection-getFirst-member(1))|Gets the first comment in the collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.commentcollection#word-word-commentcollection-getFirstOrNullObject-member(1))|Gets the first comment or null object in the collection.|
||[getItem(index: number)](/javascript/api/word/word.commentcollection#word-word-commentcollection-getItem-member(1))|Gets a comment object by its index in the collection.|
||[items](/javascript/api/word/word.commentcollection#word-word-commentcollection-items-member)|Gets the loaded child items in this collection.|
|[CommentReply](/javascript/api/word/word.commentreply)|[authorEmail](/javascript/api/word/word.commentreply#word-word-commentreply-authorEmail-member)|Gets the email of the comment reply's author.|
||[authorName](/javascript/api/word/word.commentreply#word-word-commentreply-authorName-member)|Gets the name of the comment reply's author.|
||[content](/javascript/api/word/word.commentreply#word-word-commentreply-content-member)|Gets or sets the comment reply's content.|
||[creationDate](/javascript/api/word/word.commentreply#word-word-commentreply-creationDate-member)|Gets the creation date of the comment reply.|
||[delete()](/javascript/api/word/word.commentreply#word-word-commentreply-delete-member(1))|Deletes the comment reply.|
||[id](/javascript/api/word/word.commentreply#word-word-commentreply-id-member)|ID|
||[parentComment](/javascript/api/word/word.commentreply#word-word-commentreply-parentComment-member)|Gets the parent comment of this reply.|
|[CommentReplyCollection](/javascript/api/word/word.commentreplycollection)|[getFirst()](/javascript/api/word/word.commentreplycollection#word-word-commentreplycollection-getFirst-member(1))|Gets the first comment reply in the collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.commentreplycollection#word-word-commentreplycollection-getFirstOrNullObject-member(1))|Gets the first comment reply or null object in the collection.|
||[getItem(index: number)](/javascript/api/word/word.commentreplycollection#word-word-commentreplycollection-getItem-member(1))|Gets a comment reply object by its index in the collection.|
||[items](/javascript/api/word/word.commentreplycollection#word-word-commentreplycollection-items-member)|Gets the loaded child items in this collection.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[endnotes](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-endnotes-member)|Gets the collection of endnotes in the contentcontrol.|
||[footnotes](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-footnotes-member)|Gets the collection of footnotes in the contentcontrol.|
||[getComments()](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-getComments-member(1))|Gets comments associated with the body.|
||[getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-getReviewedText-member(1))|Gets reviewed text based on ChangeTrackingVersion selection.|
|[Document](/javascript/api/word/word.document)|[changeTrackingMode](/javascript/api/word/word.document#word-word-document-changeTrackingMode-member)|Gets or sets the ChangeTracking mode.|
||[getEndnoteBody()](/javascript/api/word/word.document#word-word-document-getEndnoteBody-member(1))|Gets the document's endnotes in a single body.|
||[getFootnoteBody()](/javascript/api/word/word.document#word-word-document-getFootnoteBody-member(1))|Gets the document's footnotes in a single body.|
|[NoteItem](/javascript/api/word/word.noteitem)|[body](/javascript/api/word/word.noteitem#word-word-noteitem-body-member)|Represents the body object of the note item.|
||[delete()](/javascript/api/word/word.noteitem#word-word-noteitem-delete-member(1))|Deletes the note item.|
||[getNext()](/javascript/api/word/word.noteitem#word-word-noteitem-getNext-member(1))|Gets the next note item of the same type.|
||[getNextOrNullObject()](/javascript/api/word/word.noteitem#word-word-noteitem-getNextOrNullObject-member(1))|Gets the next note item of the same type.|
||[reference](/javascript/api/word/word.noteitem#word-word-noteitem-reference-member)|Represents a footnote or endnote reference in the main document.|
||[type](/javascript/api/word/word.noteitem#word-word-noteitem-type-member)|Represents the note item type: footnote or endnote.|
|[NoteItemCollection](/javascript/api/word/word.noteitemcollection)|[getFirst()](/javascript/api/word/word.noteitemcollection#word-word-noteitemcollection-getFirst-member(1))|Gets the first note item in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.noteitemcollection#word-word-noteitemcollection-getFirstOrNullObject-member(1))|Gets the first note item in this collection.|
||[items](/javascript/api/word/word.noteitemcollection#word-word-noteitemcollection-items-member)|Gets the loaded child items in this collection.|
|[Paragraph](/javascript/api/word/word.paragraph)|[endnotes](/javascript/api/word/word.paragraph#word-word-paragraph-endnotes-member)|Gets the collection of endnotes in the paragraph.|
||[footnotes](/javascript/api/word/word.paragraph#word-word-paragraph-footnotes-member)|Gets the collection of footnotes in the paragraph.|
||[getComments()](/javascript/api/word/word.paragraph#word-word-paragraph-getComments-member(1))|Gets comments associated with the paragraph.|
||[getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion)](/javascript/api/word/word.paragraph#word-word-paragraph-getReviewedText-member(1))|Gets reviewed text based on ChangeTrackingVersion selection.|
|[Range](/javascript/api/word/word.range)|[endnotes](/javascript/api/word/word.range#word-word-range-endnotes-member)|Gets the collection of endnotes in the range.|
||[footnotes](/javascript/api/word/word.range#word-word-range-footnotes-member)|Gets the collection of footnotes in the range.|
||[getComments()](/javascript/api/word/word.range#word-word-range-getComments-member(1))|Gets comments associated with the range.|
||[getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion)](/javascript/api/word/word.range#word-word-range-getReviewedText-member(1))|Gets reviewed text based on ChangeTrackingVersion selection.|
||[insertComment(commentText: string)](/javascript/api/word/word.range#word-word-range-insertComment-member(1))|Insert a comment on the range.|
||[insertEndnote(insertText?: string)](/javascript/api/word/word.range#word-word-range-insertEndnote-member(1))|Inserts an endnote.|
||[insertFootnote(insertText?: string)](/javascript/api/word/word.range#word-word-range-insertFootnote-member(1))|Inserts a footnote.|
|[Table](/javascript/api/word/word.table)|[endnotes](/javascript/api/word/word.table#word-word-table-endnotes-member)|Gets the collection of endnotes in the table.|
||[footnotes](/javascript/api/word/word.table#word-word-table-footnotes-member)|Gets the collection of footnotes in the table.|
|[TableRow](/javascript/api/word/word.tablerow)|[endnotes](/javascript/api/word/word.tablerow#word-word-tablerow-endnotes-member)|Gets the collection of endnotes in the table row.|
||[footnotes](/javascript/api/word/word.tablerow#word-word-tablerow-footnotes-member)|Gets the collection of footnotes in the table row.|

## See also

- [Word JavaScript API Reference Documentation](/javascript/api/word)
- [Word JavaScript API requirement sets](word-api-requirement-sets.md)
