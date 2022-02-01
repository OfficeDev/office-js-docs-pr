---
title: Word JavaScript preview APIs
description: Details about upcoming Word JavaScript APIs.
ms.date: 02/01/2022
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
|[ContentControl](/javascript/api/word/word.contentcontrol)|[onDataChanged](/javascript/api/word/word.contentcontrol#onDataChanged)|Occurs when data within the content control are changed.|
||[onDeleted](/javascript/api/word/word.contentcontrol#onDeleted)|Occurs when the content control is deleted.|
||[onSelectionChanged](/javascript/api/word/word.contentcontrol#onSelectionChanged)|Occurs when selection within the content control is changed.|
|[ContentControlEventArgs](/javascript/api/word/word.contentcontroleventargs)|[contentControl](/javascript/api/word/word.contentcontroleventargs#contentControl)|The object that raised the event.|
||[eventType](/javascript/api/word/word.contentcontroleventargs#eventType)|The event type.|
|[CustomXmlPart](/javascript/api/word/word.customxmlpart)|[delete()](/javascript/api/word/word.customxmlpart#delete__)|Deletes the custom XML part.|
||[deleteAttribute(xpath: string, namespaceMappings: any, name: string)](/javascript/api/word/word.customxmlpart#deleteAttribute_xpath__namespaceMappings__name_)|Deletes an attribute with the given name from the element identified by xpath.|
||[deleteElement(xpath: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#deleteElement_xpath__namespaceMappings_)|Deletes the element identified by xpath.|
||[getXml()](/javascript/api/word/word.customxmlpart#getXml__)|Gets the full XML content of the custom XML part.|
||[id](/javascript/api/word/word.customxmlpart#id)|Gets the ID of the custom XML part.|
||[insertAttribute(xpath: string, namespaceMappings: any, name: string, value: string)](/javascript/api/word/word.customxmlpart#insertAttribute_xpath__namespaceMappings__name__value_)|Inserts an attribute with the given name and value to the element identified by xpath.|
||[insertElement(xpath: string, xml: string, namespaceMappings: any, index?: number)](/javascript/api/word/word.customxmlpart#insertElement_xpath__xml__namespaceMappings__index_)|Inserts the given XML under the parent element identified by xpath at child position index.|
||[namespaceUri](/javascript/api/word/word.customxmlpart#namespaceUri)|Gets the namespace URI of the custom XML part.|
||[query(xpath: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#query_xpath__namespaceMappings_)|Queries the XML content of the custom XML part.|
||[setXml(xml: string)](/javascript/api/word/word.customxmlpart#setXml_xml_)|Sets the full XML content of the custom XML part.|
||[updateAttribute(xpath: string, namespaceMappings: any, name: string, value: string)](/javascript/api/word/word.customxmlpart#updateAttribute_xpath__namespaceMappings__name__value_)|Updates the value of an attribute with the given name of the element identified by xpath.|
||[updateElement(xpath: string, xml: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#updateElement_xpath__xml__namespaceMappings_)|Updates the XML of the element identified by xpath.|
|[CustomXmlPartCollection](/javascript/api/word/word.customxmlpartcollection)|[add(xml: string)](/javascript/api/word/word.customxmlpartcollection#add_xml_)|Adds a new custom XML part to the document.|
||[getByNamespace(namespaceUri: string)](/javascript/api/word/word.customxmlpartcollection#getByNamespace_namespaceUri_)|Gets a new scoped collection of custom XML parts whose namespaces match the given namespace.|
||[getCount()](/javascript/api/word/word.customxmlpartcollection#getCount__)|Gets the number of items in the collection.|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartcollection#getItem_id_)|Gets a custom XML part based on its ID.|
||[getItemOrNullObject(id: string)](/javascript/api/word/word.customxmlpartcollection#getItemOrNullObject_id_)|Gets a custom XML part based on its ID.|
||[items](/javascript/api/word/word.customxmlpartcollection#items)|Gets the loaded child items in this collection.|
|[CustomXmlPartScopedCollection](/javascript/api/word/word.customxmlpartscopedcollection)|[getCount()](/javascript/api/word/word.customxmlpartscopedcollection#getCount__)|Gets the number of items in the collection.|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartscopedcollection#getItem_id_)|Gets a custom XML part based on its ID.|
||[getItemOrNullObject(id: string)](/javascript/api/word/word.customxmlpartscopedcollection#getItemOrNullObject_id_)|Gets a custom XML part based on its ID.|
||[getOnlyItem()](/javascript/api/word/word.customxmlpartscopedcollection#getOnlyItem__)|If the collection contains exactly one item, this method returns it.|
||[getOnlyItemOrNullObject()](/javascript/api/word/word.customxmlpartscopedcollection#getOnlyItemOrNullObject__)|If the collection contains exactly one item, this method returns it.|
||[items](/javascript/api/word/word.customxmlpartscopedcollection#items)|Gets the loaded child items in this collection.|
|[Document](/javascript/api/word/word.document)|[customXmlParts](/javascript/api/word/word.document#customXmlParts)|Gets the custom XML parts in the document.|
||[deleteBookmark(name: string)](/javascript/api/word/word.document#deleteBookmark_name_)|Deletes a bookmark, if it exists, from the document.|
||[getBookmarkRange(name: string)](/javascript/api/word/word.document#getBookmarkRange_name_)|Gets a bookmark's range.|
||[getBookmarkRangeOrNullObject(name: string)](/javascript/api/word/word.document#getBookmarkRangeOrNullObject_name_)|Gets a bookmark's range.|
||[ignorePunct](/javascript/api/word/word.document#ignorePunct)||
||[ignoreSpace](/javascript/api/word/word.document#ignoreSpace)||
||[matchCase](/javascript/api/word/word.document#matchCase)||
||[matchPrefix](/javascript/api/word/word.document#matchPrefix)||
||[matchSuffix](/javascript/api/word/word.document#matchSuffix)||
||[matchWholeWord](/javascript/api/word/word.document#matchWholeWord)||
||[matchWildcards](/javascript/api/word/word.document#matchWildcards)||
||[onContentControlAdded](/javascript/api/word/word.document#onContentControlAdded)|Occurs when a content control is added.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| {            ignorePunct?: boolean            ignoreSpace?: boolean            matchCase?: boolean            matchPrefix?: boolean            matchSuffix?: boolean            matchWholeWord?: boolean            matchWildcards?: boolean        })](/javascript/api/word/word.document#search_searchText__searchOptions_)|Performs a search with the specified search options on the scope of the whole document.|
||[settings](/javascript/api/word/word.document#settings)|Gets the add-in's settings in the document.|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[customXmlParts](/javascript/api/word/word.documentcreated#customXmlParts)|Gets the custom XML parts in the document.|
||[deleteBookmark(name: string)](/javascript/api/word/word.documentcreated#deleteBookmark_name_)|Deletes a bookmark, if it exists, from the document.|
||[getBookmarkRange(name: string)](/javascript/api/word/word.documentcreated#getBookmarkRange_name_)|Gets a bookmark's range.|
||[getBookmarkRangeOrNullObject(name: string)](/javascript/api/word/word.documentcreated#getBookmarkRangeOrNullObject_name_)|Gets a bookmark's range.|
||[settings](/javascript/api/word/word.documentcreated#settings)|Gets the add-in's settings in the document.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[imageFormat](/javascript/api/word/word.inlinepicture#imageFormat)|Gets the format of the inline image.|
|[List](/javascript/api/word/word.list)|[getLevelFont(level: number)](/javascript/api/word/word.list#getLevelFont_level_)|Gets the font of the bullet, number, or picture at the specified level in the list.|
||[getLevelPicture(level: number)](/javascript/api/word/word.list#getLevelPicture_level_)|Gets the base64 encoded string representation of the picture at the specified level in the list.|
||[resetLevelFont(level: number, resetFontName?: boolean)](/javascript/api/word/word.list#resetLevelFont_level__resetFontName_)|Resets the font of the bullet, number, or picture at the specified level in the list.|
||[setLevelPicture(level: number, base64EncodedImage?: string)](/javascript/api/word/word.list#setLevelPicture_level__base64EncodedImage_)|Sets the picture at the specified level in the list.|
|[Range](/javascript/api/word/word.range)|[getBookmarks(includeHidden?: boolean, includeAdjacent?: boolean)](/javascript/api/word/word.range#getBookmarks_includeHidden__includeAdjacent_)|Gets the names all bookmarks in or overlapping the range.|
||[insertBookmark(name: string)](/javascript/api/word/word.range#insertBookmark_name_)|Inserts a bookmark on the range.|
|[Setting](/javascript/api/word/word.setting)|[delete()](/javascript/api/word/word.setting#delete__)|Deletes the setting.|
||[key](/javascript/api/word/word.setting#key)|Gets the key of the setting.|
||[value](/javascript/api/word/word.setting#value)|Gets or sets the value of the setting.|
|[SettingCollection](/javascript/api/word/word.settingcollection)|[add(key: string, value: any)](/javascript/api/word/word.settingcollection#add_key__value_)|Creates a new setting or sets an existing setting.|
||[deleteAll()](/javascript/api/word/word.settingcollection#deleteAll__)|Deletes all settings in this add-in.|
||[getCount()](/javascript/api/word/word.settingcollection#getCount__)|Gets the count of settings.|
||[getItem(key: string)](/javascript/api/word/word.settingcollection#getItem_key_)|Gets a setting object by its key, which is case-sensitive.|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.settingcollection#getItemOrNullObject_key_)|Gets a setting object by its key, which is case-sensitive.|
||[items](/javascript/api/word/word.settingcollection#items)|Gets the loaded child items in this collection.|
|[Table](/javascript/api/word/word.table)|[mergeCells(topRow: number, firstCell: number, bottomRow: number, lastCell: number)](/javascript/api/word/word.table#mergeCells_topRow__firstCell__bottomRow__lastCell_)|Merges the cells bounded inclusively by a first and last cell.|
|[TableCell](/javascript/api/word/word.tablecell)|[split(rowCount: number, columnCount: number)](/javascript/api/word/word.tablecell#split_rowCount__columnCount_)|Splits the cell into the specified number of rows and columns.|
|[TableRow](/javascript/api/word/word.tablerow)|[insertContentControl()](/javascript/api/word/word.tablerow#insertContentControl__)|Inserts a content control on the row.|
||[merge()](/javascript/api/word/word.tablerow#merge__)|Merges the row into one cell.|

## Web-only API list

The following table lists the Word JavaScript APIs currently in preview only in Word on the web. To see a complete list of all Word JavaScript APIs (including preview APIs and previously released APIs), see [all Word JavaScript APIs](/javascript/api/word?view=word-js-preview&preserve-view=true).

| Class | Fields | Description |
|:---|:---|:---|
|[Body](/javascript/api/word/word.body)|[endnotes](/javascript/api/word/word.body#endnotes)|Gets the collection of endnotes in the body.|
||[footnotes](/javascript/api/word/word.body#footnotes)|Gets the collection of footnotes in the body.|
||[getComments()](/javascript/api/word/word.body#getComments__)|Gets comments associated with the body.|
||[getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion)](/javascript/api/word/word.body#getReviewedText_changeTrackingVersion_)|Gets reviewed text based on ChangeTrackingVersion selection.|
||[type](/javascript/api/word/word.body#type)|Gets the type of the body.|
|[Comment](/javascript/api/word/word.comment)|[authorEmail](/javascript/api/word/word.comment#authorEmail)|Gets the email of the comment's author.|
||[authorName](/javascript/api/word/word.comment#authorName)|Gets the name of the comment's author.|
||[content](/javascript/api/word/word.comment#content)|Gets or sets the comment's content as plain text.|
||[contentRange](/javascript/api/word/word.comment#contentRange)|Gets or sets the comment thread status.|
||[creationDate](/javascript/api/word/word.comment#creationDate)|Gets the creation date of the comment.|
||[delete()](/javascript/api/word/word.comment#delete__)|Deletes the comment and its replies.|
||[getRange()](/javascript/api/word/word.comment#getRange__)|Gets the range in the main document where the comment is on.|
||[id](/javascript/api/word/word.comment#id)|ID|
||[replies](/javascript/api/word/word.comment#replies)|Gets the collection of reply objects associated with the comment.|
||[reply(replyText: string)](/javascript/api/word/word.comment#reply_replyText_)|Adds a new reply to the end of the comment thread.|
||[resolved](/javascript/api/word/word.comment#resolved)|Gets or sets the comment thread's status.|
|[CommentCollection](/javascript/api/word/word.commentcollection)|[getFirst()](/javascript/api/word/word.commentcollection#getFirst__)|Gets the first comment in the collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.commentcollection#getFirstOrNullObject__)|Gets the first comment in the collection.|
||[getItem(index: number)](/javascript/api/word/word.commentcollection#getItem_index_)|Gets a comment object by its index in the collection.|
||[items](/javascript/api/word/word.commentcollection#items)|Gets the loaded child items in this collection.|
|[CommentContentRange](/javascript/api/word/word.commentcontentrange)|[bold](/javascript/api/word/word.commentcontentrange#bold)|Gets or sets a value that indicates whether the comment text is bold.|
||[hyperlink](/javascript/api/word/word.commentcontentrange#hyperlink)|Gets the first hyperlink in the range, or sets a hyperlink on the range.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.commentcontentrange#insertText_text__insertLocation_)|Inserts text into at the specified location.|
||[isEmpty](/javascript/api/word/word.commentcontentrange#isEmpty)|Checks whether the range length is zero.|
||[italic](/javascript/api/word/word.commentcontentrange#italic)|Gets or sets a value that indicates whether the comment text is italicized.|
||[strikeThrough](/javascript/api/word/word.commentcontentrange#strikeThrough)|Gets or sets a value that indicates whether the comment text has a strikethrough.|
||[text](/javascript/api/word/word.commentcontentrange#text)|Gets the text of the comment range.|
||[underline](/javascript/api/word/word.commentcontentrange#underline)|Gets or sets a value that indicates the comment text's underline type.|
|[CommentReply](/javascript/api/word/word.commentreply)|[authorEmail](/javascript/api/word/word.commentreply#authorEmail)|Gets the email of the comment reply's author.|
||[authorName](/javascript/api/word/word.commentreply#authorName)|Gets the name of the comment reply's author.|
||[content](/javascript/api/word/word.commentreply#content)|Gets or sets the comment reply's content.|
||[contentRange](/javascript/api/word/word.commentreply#contentRange)|Gets or sets the commentReply's content range.|
||[creationDate](/javascript/api/word/word.commentreply#creationDate)|Gets the creation date of the comment reply.|
||[delete()](/javascript/api/word/word.commentreply#delete__)|Deletes the comment reply.|
||[id](/javascript/api/word/word.commentreply#id)|ID|
||[parentComment](/javascript/api/word/word.commentreply#parentComment)|Gets the parent comment of this reply.|
|[CommentReplyCollection](/javascript/api/word/word.commentreplycollection)|[getFirst()](/javascript/api/word/word.commentreplycollection#getFirst__)|Gets the first comment reply in the collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.commentreplycollection#getFirstOrNullObject__)|Gets the first comment reply in the collection.|
||[getItem(index: number)](/javascript/api/word/word.commentreplycollection#getItem_index_)|Gets a comment reply object by its index in the collection.|
||[items](/javascript/api/word/word.commentreplycollection#items)|Gets the loaded child items in this collection.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[endnotes](/javascript/api/word/word.contentcontrol#endnotes)|Gets the collection of endnotes in the contentcontrol.|
||[footnotes](/javascript/api/word/word.contentcontrol#footnotes)|Gets the collection of footnotes in the contentcontrol.|
||[getComments()](/javascript/api/word/word.contentcontrol#getComments__)|Gets comments associated with the body.|
||[getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion)](/javascript/api/word/word.contentcontrol#getReviewedText_changeTrackingVersion_)|Gets reviewed text based on ChangeTrackingVersion selection.|
|[Document](/javascript/api/word/word.document)|[changeTrackingMode](/javascript/api/word/word.document#changeTrackingMode)|Gets or sets the ChangeTracking mode.|
||[getEndnoteBody()](/javascript/api/word/word.document#getEndnoteBody__)|Gets the document's endnotes in a single body.|
||[getFootnoteBody()](/javascript/api/word/word.document#getFootnoteBody__)|Gets the document's footnotes in a single body.|
|[NoteItem](/javascript/api/word/word.noteitem)|[body](/javascript/api/word/word.noteitem#body)|Represents the body object of the note item.|
||[delete()](/javascript/api/word/word.noteitem#delete__)|Deletes the note item.|
||[getNext()](/javascript/api/word/word.noteitem#getNext__)|Gets the next note item of the same type.|
||[getNextOrNullObject()](/javascript/api/word/word.noteitem#getNextOrNullObject__)|Gets the next note item of the same type.|
||[reference](/javascript/api/word/word.noteitem#reference)|Represents a footnote or endnote reference in the main document.|
||[type](/javascript/api/word/word.noteitem#type)|Represents the note item type: footnote or endnote.|
|[NoteItemCollection](/javascript/api/word/word.noteitemcollection)|[getFirst()](/javascript/api/word/word.noteitemcollection#getFirst__)|Gets the first note item in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.noteitemcollection#getFirstOrNullObject__)|Gets the first note item in this collection.|
||[items](/javascript/api/word/word.noteitemcollection#items)|Gets the loaded child items in this collection.|
|[Paragraph](/javascript/api/word/word.paragraph)|[endnotes](/javascript/api/word/word.paragraph#endnotes)|Gets the collection of endnotes in the paragraph.|
||[footnotes](/javascript/api/word/word.paragraph#footnotes)|Gets the collection of footnotes in the paragraph.|
||[getComments()](/javascript/api/word/word.paragraph#getComments__)|Gets comments associated with the paragraph.|
||[getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion)](/javascript/api/word/word.paragraph#getReviewedText_changeTrackingVersion_)|Gets reviewed text based on ChangeTrackingVersion selection.|
|[Range](/javascript/api/word/word.range)|[endnotes](/javascript/api/word/word.range#endnotes)|Gets the collection of endnotes in the range.|
||[footnotes](/javascript/api/word/word.range#footnotes)|Gets the collection of footnotes in the range.|
||[getComments()](/javascript/api/word/word.range#getComments__)|Gets comments associated with the range.|
||[getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion)](/javascript/api/word/word.range#getReviewedText_changeTrackingVersion_)|Gets reviewed text based on ChangeTrackingVersion selection.|
||[insertComment(commentText: string)](/javascript/api/word/word.range#insertComment_commentText_)|Insert a comment on the range.|
||[insertEndnote(insertText?: string)](/javascript/api/word/word.range#insertEndnote_insertText_)|Inserts an endnote.|
||[insertFootnote(insertText?: string)](/javascript/api/word/word.range#insertFootnote_insertText_)|Inserts a footnote.|
|[Table](/javascript/api/word/word.table)|[endnotes](/javascript/api/word/word.table#endnotes)|Gets the collection of endnotes in the table.|
||[footnotes](/javascript/api/word/word.table#footnotes)|Gets the collection of footnotes in the table.|
|[TableRow](/javascript/api/word/word.tablerow)|[endnotes](/javascript/api/word/word.tablerow#endnotes)|Gets the collection of endnotes in the table row.|
||[footnotes](/javascript/api/word/word.tablerow#footnotes)|Gets the collection of footnotes in the table row.|

## See also

- [Word JavaScript API Reference Documentation](/javascript/api/word)
- [Word JavaScript API requirement sets](word-api-requirement-sets.md)
