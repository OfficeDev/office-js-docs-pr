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
|[ContentControl](/javascript/api/word/word.contentcontrol)|[onDataChanged](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-ondatachanged-member)|Occurs when data within the content control are changed.|
||[onDeleted](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-ondeleted-member)|Occurs when the content control is deleted.|
||[onSelectionChanged](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-onselectionchanged-member)|Occurs when selection within the content control is changed.|
|[ContentControlEventArgs](/javascript/api/word/word.contentcontroleventargs)|[contentControl](/javascript/api/word/word.contentcontroleventargs#word-word-contentcontroleventargs-contentcontrol-member)|The object that raised the event.|
||[eventType](/javascript/api/word/word.contentcontroleventargs#word-word-contentcontroleventargs-eventtype-member)|The event type.|
|[CustomXmlPart](/javascript/api/word/word.customxmlpart)|[delete()](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-delete-member(1))|Deletes the custom XML part.|
||[deleteAttribute(xpath: string, namespaceMappings: any, name: string)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-deleteattribute-member(1))|Deletes an attribute with the given name from the element identified by xpath.|
||[deleteElement(xpath: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-deleteelement-member(1))|Deletes the element identified by xpath.|
||[getXml()](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-getxml-member(1))|Gets the full XML content of the custom XML part.|
||[id](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-id-member)|Gets the ID of the custom XML part.|
||[insertAttribute(xpath: string, namespaceMappings: any, name: string, value: string)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-insertattribute-member(1))|Inserts an attribute with the given name and value to the element identified by xpath.|
||[insertElement(xpath: string, xml: string, namespaceMappings: any, index?: number)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-insertelement-member(1))|Inserts the given XML under the parent element identified by xpath at child position index.|
||[namespaceUri](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-namespaceuri-member)|Gets the namespace URI of the custom XML part.|
||[query(xpath: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-query-member(1))|Queries the XML content of the custom XML part.|
||[setXml(xml: string)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-setxml-member(1))|Sets the full XML content of the custom XML part.|
||[updateAttribute(xpath: string, namespaceMappings: any, name: string, value: string)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-updateattribute-member(1))|Updates the value of an attribute with the given name of the element identified by xpath.|
||[updateElement(xpath: string, xml: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-updateelement-member(1))|Updates the XML of the element identified by xpath.|
|[CustomXmlPartCollection](/javascript/api/word/word.customxmlpartcollection)|[add(xml: string)](/javascript/api/word/word.customxmlpartcollection#word-word-customxmlpartcollection-add-member(1))|Adds a new custom XML part to the document.|
||[getByNamespace(namespaceUri: string)](/javascript/api/word/word.customxmlpartcollection#word-word-customxmlpartcollection-getbynamespace-member(1))|Gets a new scoped collection of custom XML parts whose namespaces match the given namespace.|
||[getCount()](/javascript/api/word/word.customxmlpartcollection#word-word-customxmlpartcollection-getcount-member(1))|Gets the number of items in the collection.|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartcollection#word-word-customxmlpartcollection-getitem-member(1))|Gets a custom XML part based on its ID.|
||[getItemOrNullObject(id: string)](/javascript/api/word/word.customxmlpartcollection#word-word-customxmlpartcollection-getitemornullobject-member(1))|Gets a custom XML part based on its ID.|
||[items](/javascript/api/word/word.customxmlpartcollection#word-word-customxmlpartcollection-items-member)|Gets the loaded child items in this collection.|
|[CustomXmlPartScopedCollection](/javascript/api/word/word.customxmlpartscopedcollection)|[getCount()](/javascript/api/word/word.customxmlpartscopedcollection#word-word-customxmlpartscopedcollection-getcount-member(1))|Gets the number of items in the collection.|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartscopedcollection#word-word-customxmlpartscopedcollection-getitem-member(1))|Gets a custom XML part based on its ID.|
||[getItemOrNullObject(id: string)](/javascript/api/word/word.customxmlpartscopedcollection#word-word-customxmlpartscopedcollection-getitemornullobject-member(1))|Gets a custom XML part based on its ID.|
||[getOnlyItem()](/javascript/api/word/word.customxmlpartscopedcollection#word-word-customxmlpartscopedcollection-getonlyitem-member(1))|If the collection contains exactly one item, this method returns it.|
||[getOnlyItemOrNullObject()](/javascript/api/word/word.customxmlpartscopedcollection#word-word-customxmlpartscopedcollection-getonlyitemornullobject-member(1))|If the collection contains exactly one item, this method returns it.|
||[items](/javascript/api/word/word.customxmlpartscopedcollection#word-word-customxmlpartscopedcollection-items-member)|Gets the loaded child items in this collection.|
|[Document](/javascript/api/word/word.document)|[customXmlParts](/javascript/api/word/word.document#word-word-document-customxmlparts-member)|Gets the custom XML parts in the document.|
||[deleteBookmark(name: string)](/javascript/api/word/word.document#word-word-document-deletebookmark-member(1))|Deletes a bookmark, if it exists, from the document.|
||[getBookmarkRange(name: string)](/javascript/api/word/word.document#word-word-document-getbookmarkrange-member(1))|Gets a bookmark's range.|
||[getBookmarkRangeOrNullObject(name: string)](/javascript/api/word/word.document#word-word-document-getbookmarkrangeornullobject-member(1))|Gets a bookmark's range.|
||[ignorePunct](/javascript/api/word/word.document#word-word-document-ignorepunct-member)||
||[ignoreSpace](/javascript/api/word/word.document#word-word-document-ignorespace-member)||
||[matchCase](/javascript/api/word/word.document#word-word-document-matchcase-member)||
||[matchPrefix](/javascript/api/word/word.document#word-word-document-matchprefix-member)||
||[matchSuffix](/javascript/api/word/word.document#word-word-document-matchsuffix-member)||
||[matchWholeWord](/javascript/api/word/word.document#word-word-document-matchwholeword-member)||
||[matchWildcards](/javascript/api/word/word.document#word-word-document-matchwildcards-member)||
||[onContentControlAdded](/javascript/api/word/word.document#word-word-document-oncontentcontroladded-member)|Occurs when a content control is added.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| {            ignorePunct?: boolean            ignoreSpace?: boolean            matchCase?: boolean            matchPrefix?: boolean            matchSuffix?: boolean            matchWholeWord?: boolean            matchWildcards?: boolean        })](/javascript/api/word/word.document#word-word-document-search-member(1))|Performs a search with the specified search options on the scope of the whole document.|
||[settings](/javascript/api/word/word.document#word-word-document-settings-member)|Gets the add-in's settings in the document.|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[customXmlParts](/javascript/api/word/word.documentcreated#word-word-documentcreated-customxmlparts-member)|Gets the custom XML parts in the document.|
||[deleteBookmark(name: string)](/javascript/api/word/word.documentcreated#word-word-documentcreated-deletebookmark-member(1))|Deletes a bookmark, if it exists, from the document.|
||[getBookmarkRange(name: string)](/javascript/api/word/word.documentcreated#word-word-documentcreated-getbookmarkrange-member(1))|Gets a bookmark's range.|
||[getBookmarkRangeOrNullObject(name: string)](/javascript/api/word/word.documentcreated#word-word-documentcreated-getbookmarkrangeornullobject-member(1))|Gets a bookmark's range.|
||[settings](/javascript/api/word/word.documentcreated#word-word-documentcreated-settings-member)|Gets the add-in's settings in the document.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[imageFormat](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-imageformat-member)|Gets the format of the inline image.|
|[List](/javascript/api/word/word.list)|[getLevelFont(level: number)](/javascript/api/word/word.list#word-word-list-getlevelfont-member(1))|Gets the font of the bullet, number, or picture at the specified level in the list.|
||[getLevelPicture(level: number)](/javascript/api/word/word.list#word-word-list-getlevelpicture-member(1))|Gets the base64 encoded string representation of the picture at the specified level in the list.|
||[resetLevelFont(level: number, resetFontName?: boolean)](/javascript/api/word/word.list#word-word-list-resetlevelfont-member(1))|Resets the font of the bullet, number, or picture at the specified level in the list.|
||[setLevelPicture(level: number, base64EncodedImage?: string)](/javascript/api/word/word.list#word-word-list-setlevelpicture-member(1))|Sets the picture at the specified level in the list.|
|[Range](/javascript/api/word/word.range)|[getBookmarks(includeHidden?: boolean, includeAdjacent?: boolean)](/javascript/api/word/word.range#word-word-range-getbookmarks-member(1))|Gets the names all bookmarks in or overlapping the range.|
||[insertBookmark(name: string)](/javascript/api/word/word.range#word-word-range-insertbookmark-member(1))|Inserts a bookmark on the range.|
|[Setting](/javascript/api/word/word.setting)|[delete()](/javascript/api/word/word.setting#word-word-setting-delete-member(1))|Deletes the setting.|
||[key](/javascript/api/word/word.setting#word-word-setting-key-member)|Gets the key of the setting.|
||[value](/javascript/api/word/word.setting#word-word-setting-value-member)|Gets or sets the value of the setting.|
|[SettingCollection](/javascript/api/word/word.settingcollection)|[add(key: string, value: any)](/javascript/api/word/word.settingcollection#word-word-settingcollection-add-member(1))|Creates a new setting or sets an existing setting.|
||[deleteAll()](/javascript/api/word/word.settingcollection#word-word-settingcollection-deleteall-member(1))|Deletes all settings in this add-in.|
||[getCount()](/javascript/api/word/word.settingcollection#word-word-settingcollection-getcount-member(1))|Gets the count of settings.|
||[getItem(key: string)](/javascript/api/word/word.settingcollection#word-word-settingcollection-getitem-member(1))|Gets a setting object by its key, which is case-sensitive.|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.settingcollection#word-word-settingcollection-getitemornullobject-member(1))|Gets a setting object by its key, which is case-sensitive.|
||[items](/javascript/api/word/word.settingcollection#word-word-settingcollection-items-member)|Gets the loaded child items in this collection.|
|[Table](/javascript/api/word/word.table)|[mergeCells(topRow: number, firstCell: number, bottomRow: number, lastCell: number)](/javascript/api/word/word.table#word-word-table-mergecells-member(1))|Merges the cells bounded inclusively by a first and last cell.|
|[TableCell](/javascript/api/word/word.tablecell)|[split(rowCount: number, columnCount: number)](/javascript/api/word/word.tablecell#word-word-tablecell-split-member(1))|Splits the cell into the specified number of rows and columns.|
|[TableRow](/javascript/api/word/word.tablerow)|[insertContentControl()](/javascript/api/word/word.tablerow#word-word-tablerow-insertcontentcontrol-member(1))|Inserts a content control on the row.|
||[merge()](/javascript/api/word/word.tablerow#word-word-tablerow-merge-member(1))|Merges the row into one cell.|

## Web-only API list

The following table lists the Word JavaScript APIs currently in preview only in Word on the web. To see a complete list of all Word JavaScript APIs (including preview APIs and previously released APIs), see [all Word JavaScript APIs](/javascript/api/word?view=word-js-preview&preserve-view=true).

| Class | Fields | Description |
|:---|:---|:---|
|[Body](/javascript/api/word/word.body)|[endnotes](/javascript/api/word/word.body#word-word-body-endnotes-member)|Gets the collection of endnotes in the body.|
||[footnotes](/javascript/api/word/word.body#word-word-body-footnotes-member)|Gets the collection of footnotes in the body.|
||[getComments()](/javascript/api/word/word.body#word-word-body-getcomments-member(1))|Gets comments associated with the body.|
||[getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion)](/javascript/api/word/word.body#word-word-body-getreviewedtext-member(1))|Gets reviewed text based on ChangeTrackingVersion selection.|
||[type](/javascript/api/word/word.body#word-word-body-type-member)|Gets the type of the body.|
|[Comment](/javascript/api/word/word.comment)|[authorEmail](/javascript/api/word/word.comment#word-word-comment-authoremail-member)|Gets the email of the comment's author.|
||[authorName](/javascript/api/word/word.comment#word-word-comment-authorname-member)|Gets the name of the comment's author.|
||[content](/javascript/api/word/word.comment#word-word-comment-content-member)|Gets or sets the comment's content as plain text.|
||[contentRange](/javascript/api/word/word.comment#word-word-comment-contentrange-member)|Gets or sets the comment thread status.|
||[creationDate](/javascript/api/word/word.comment#word-word-comment-creationdate-member)|Gets the creation date of the comment.|
||[delete()](/javascript/api/word/word.comment#word-word-comment-delete-member(1))|Deletes the comment and its replies.|
||[getRange()](/javascript/api/word/word.comment#word-word-comment-getrange-member(1))|Gets the range in the main document where the comment is on.|
||[id](/javascript/api/word/word.comment#word-word-comment-id-member)|ID|
||[replies](/javascript/api/word/word.comment#word-word-comment-replies-member)|Gets the collection of reply objects associated with the comment.|
||[reply(replyText: string)](/javascript/api/word/word.comment#word-word-comment-reply-member(1))|Adds a new reply to the end of the comment thread.|
||[resolved](/javascript/api/word/word.comment#word-word-comment-resolved-member)|Gets or sets the comment thread's status.|
|[CommentCollection](/javascript/api/word/word.commentcollection)|[getFirst()](/javascript/api/word/word.commentcollection#word-word-commentcollection-getfirst-member(1))|Gets the first comment in the collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.commentcollection#word-word-commentcollection-getfirstornullobject-member(1))|Gets the first comment in the collection.|
||[getItem(index: number)](/javascript/api/word/word.commentcollection#word-word-commentcollection-getitem-member(1))|Gets a comment object by its index in the collection.|
||[items](/javascript/api/word/word.commentcollection#word-word-commentcollection-items-member)|Gets the loaded child items in this collection.|
|[CommentContentRange](/javascript/api/word/word.commentcontentrange)|[bold](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-bold-member)|Gets or sets a value that indicates whether the comment text is bold.|
||[hyperlink](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-hyperlink-member)|Gets the first hyperlink in the range, or sets a hyperlink on the range.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-inserttext-member(1))|Inserts text into at the specified location.|
||[isEmpty](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-isempty-member)|Checks whether the range length is zero.|
||[italic](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-italic-member)|Gets or sets a value that indicates whether the comment text is italicized.|
||[strikeThrough](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-strikethrough-member)|Gets or sets a value that indicates whether the comment text has a strikethrough.|
||[text](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-text-member)|Gets the text of the comment range.|
||[underline](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-underline-member)|Gets or sets a value that indicates the comment text's underline type.|
|[CommentReply](/javascript/api/word/word.commentreply)|[authorEmail](/javascript/api/word/word.commentreply#word-word-commentreply-authoremail-member)|Gets the email of the comment reply's author.|
||[authorName](/javascript/api/word/word.commentreply#word-word-commentreply-authorname-member)|Gets the name of the comment reply's author.|
||[content](/javascript/api/word/word.commentreply#word-word-commentreply-content-member)|Gets or sets the comment reply's content.|
||[contentRange](/javascript/api/word/word.commentreply#word-word-commentreply-contentrange-member)|Gets or sets the commentReply's content range.|
||[creationDate](/javascript/api/word/word.commentreply#word-word-commentreply-creationdate-member)|Gets the creation date of the comment reply.|
||[delete()](/javascript/api/word/word.commentreply#word-word-commentreply-delete-member(1))|Deletes the comment reply.|
||[id](/javascript/api/word/word.commentreply#word-word-commentreply-id-member)|ID|
||[parentComment](/javascript/api/word/word.commentreply#word-word-commentreply-parentcomment-member)|Gets the parent comment of this reply.|
|[CommentReplyCollection](/javascript/api/word/word.commentreplycollection)|[getFirst()](/javascript/api/word/word.commentreplycollection#word-word-commentreplycollection-getfirst-member(1))|Gets the first comment reply in the collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.commentreplycollection#word-word-commentreplycollection-getfirstornullobject-member(1))|Gets the first comment reply in the collection.|
||[getItem(index: number)](/javascript/api/word/word.commentreplycollection#word-word-commentreplycollection-getitem-member(1))|Gets a comment reply object by its index in the collection.|
||[items](/javascript/api/word/word.commentreplycollection#word-word-commentreplycollection-items-member)|Gets the loaded child items in this collection.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[endnotes](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-endnotes-member)|Gets the collection of endnotes in the contentcontrol.|
||[footnotes](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-footnotes-member)|Gets the collection of footnotes in the contentcontrol.|
||[getComments()](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-getcomments-member(1))|Gets comments associated with the body.|
||[getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-getreviewedtext-member(1))|Gets reviewed text based on ChangeTrackingVersion selection.|
|[Document](/javascript/api/word/word.document)|[changeTrackingMode](/javascript/api/word/word.document#word-word-document-changetrackingmode-member)|Gets or sets the ChangeTracking mode.|
||[getEndnoteBody()](/javascript/api/word/word.document#word-word-document-getendnotebody-member(1))|Gets the document's endnotes in a single body.|
||[getFootnoteBody()](/javascript/api/word/word.document#word-word-document-getfootnotebody-member(1))|Gets the document's footnotes in a single body.|
|[NoteItem](/javascript/api/word/word.noteitem)|[body](/javascript/api/word/word.noteitem#word-word-noteitem-body-member)|Represents the body object of the note item.|
||[delete()](/javascript/api/word/word.noteitem#word-word-noteitem-delete-member(1))|Deletes the note item.|
||[getNext()](/javascript/api/word/word.noteitem#word-word-noteitem-getnext-member(1))|Gets the next note item of the same type.|
||[getNextOrNullObject()](/javascript/api/word/word.noteitem#word-word-noteitem-getnextornullobject-member(1))|Gets the next note item of the same type.|
||[reference](/javascript/api/word/word.noteitem#word-word-noteitem-reference-member)|Represents a footnote or endnote reference in the main document.|
||[type](/javascript/api/word/word.noteitem#word-word-noteitem-type-member)|Represents the note item type: footnote or endnote.|
|[NoteItemCollection](/javascript/api/word/word.noteitemcollection)|[getFirst()](/javascript/api/word/word.noteitemcollection#word-word-noteitemcollection-getfirst-member(1))|Gets the first note item in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.noteitemcollection#word-word-noteitemcollection-getfirstornullobject-member(1))|Gets the first note item in this collection.|
||[items](/javascript/api/word/word.noteitemcollection#word-word-noteitemcollection-items-member)|Gets the loaded child items in this collection.|
|[Paragraph](/javascript/api/word/word.paragraph)|[endnotes](/javascript/api/word/word.paragraph#word-word-paragraph-endnotes-member)|Gets the collection of endnotes in the paragraph.|
||[footnotes](/javascript/api/word/word.paragraph#word-word-paragraph-footnotes-member)|Gets the collection of footnotes in the paragraph.|
||[getComments()](/javascript/api/word/word.paragraph#word-word-paragraph-getcomments-member(1))|Gets comments associated with the paragraph.|
||[getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion)](/javascript/api/word/word.paragraph#word-word-paragraph-getreviewedtext-member(1))|Gets reviewed text based on ChangeTrackingVersion selection.|
|[Range](/javascript/api/word/word.range)|[endnotes](/javascript/api/word/word.range#word-word-range-endnotes-member)|Gets the collection of endnotes in the range.|
||[footnotes](/javascript/api/word/word.range#word-word-range-footnotes-member)|Gets the collection of footnotes in the range.|
||[getComments()](/javascript/api/word/word.range#word-word-range-getcomments-member(1))|Gets comments associated with the range.|
||[getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion)](/javascript/api/word/word.range#word-word-range-getreviewedtext-member(1))|Gets reviewed text based on ChangeTrackingVersion selection.|
||[insertComment(commentText: string)](/javascript/api/word/word.range#word-word-range-insertcomment-member(1))|Insert a comment on the range.|
||[insertEndnote(insertText?: string)](/javascript/api/word/word.range#word-word-range-insertendnote-member(1))|Inserts an endnote.|
||[insertFootnote(insertText?: string)](/javascript/api/word/word.range#word-word-range-insertfootnote-member(1))|Inserts a footnote.|
|[Table](/javascript/api/word/word.table)|[endnotes](/javascript/api/word/word.table#word-word-table-endnotes-member)|Gets the collection of endnotes in the table.|
||[footnotes](/javascript/api/word/word.table#word-word-table-footnotes-member)|Gets the collection of footnotes in the table.|
|[TableRow](/javascript/api/word/word.tablerow)|[endnotes](/javascript/api/word/word.tablerow#word-word-tablerow-endnotes-member)|Gets the collection of endnotes in the table row.|
||[footnotes](/javascript/api/word/word.tablerow#word-word-tablerow-footnotes-member)|Gets the collection of footnotes in the table row.|

## See also

- [Word JavaScript API Reference Documentation](/javascript/api/word)
- [Word JavaScript API requirement sets](word-api-requirement-sets.md)
