---
title: Word JavaScript preview APIs
description: 'Details about upcoming Word JavaScript APIs'
ms.date: 07/17/2019
ms.prod: word
localization_priority: Normal
---

# Word JavaScript preview APIs

New Word JavaScript APIs are first introduced in "preview" and later become part of a specific, numbered requirement set after sufficient testing occurs and user feedback is acquired.

> [!NOTE]
> Preview APIs are subject to change and are not intended for use in a production environment. We recommend that you try them out in test and development environments only. Do not use preview APIs in a production environment or within business-critical documents.
>
> To use preview APIs, you must reference the **beta** library on the CDN (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) and you may also need to join the Office Insider program to get a recent Office build.

## API List

The following is a complete list of APIs currently in preview.

| Class | Fields | Description |
|:---|:---|:---|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[onDataChanged](/javascript/api/word/word.contentcontrol#ondatachanged)|Occurs when data within the content control are changed. To get the new text, load this content control in the handler. To get the old text, do not load it.|
||[onDeleted](/javascript/api/word/word.contentcontrol#ondeleted)|Occurs when the content control is deleted. Do not load this content control in the handler, otherwise you won't be able to get its original properties.|
||[onSelectionChanged](/javascript/api/word/word.contentcontrol#onselectionchanged)|Occurs when selection within the content control is changed.|
|[ContentControlEventArgs](/javascript/api/word/word.contentcontroleventargs)|[contentControl](/javascript/api/word/word.contentcontroleventargs#contentcontrol)|The object that raised the event. Load this object to get its properties.|
||[eventType](/javascript/api/word/word.contentcontroleventargs#eventtype)|The event type. See Word.EventType for details.|
|[CustomXmlPart](/javascript/api/word/word.customxmlpart)|[delete()](/javascript/api/word/word.customxmlpart#delete--)|Deletes the custom XML part.|
||[deleteAttribute(xpath: string, namespaceMappings: any, name: string)](/javascript/api/word/word.customxmlpart#deleteattribute-xpath--namespacemappings--name-)|Deletes an attribute with the given name from the element identified by xpath.|
||[deleteElement(xpath: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#deleteelement-xpath--namespacemappings-)|Deletes the element identified by xpath.|
||[getXml()](/javascript/api/word/word.customxmlpart#getxml--)|Gets the full XML content of the custom XML part.|
||[insertAttribute(xpath: string, namespaceMappings: any, name: string, value: string)](/javascript/api/word/word.customxmlpart#insertattribute-xpath--namespacemappings--name--value-)|Inserts an attribute with the given name and value to the element identified by xpath.|
||[insertElement(xpath: string, xml: string, namespaceMappings: any, index?: number)](/javascript/api/word/word.customxmlpart#insertelement-xpath--xml--namespacemappings--index-)|Inserts the given XML under the parent element identified by xpath at child position index.|
||[query(xpath: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#query-xpath--namespacemappings-)|Queries the XML content of the custom XML part.|
||[id](/javascript/api/word/word.customxmlpart#id)|Gets the ID of the custom XML part. Read only.|
||[namespaceUri](/javascript/api/word/word.customxmlpart#namespaceuri)|Gets the namespace URI of the custom XML part. Read only.|
||[setXml(xml: string)](/javascript/api/word/word.customxmlpart#setxml-xml-)|Sets the full XML content of the custom XML part.|
||[updateAttribute(xpath: string, namespaceMappings: any, name: string, value: string)](/javascript/api/word/word.customxmlpart#updateattribute-xpath--namespacemappings--name--value-)|Updates the value of an attribute with the given name of the element identified by xpath.|
||[updateElement(xpath: string, xml: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#updateelement-xpath--xml--namespacemappings-)|Updates the XML of the element identified by xpath.|
|[CustomXmlPartCollection](/javascript/api/word/word.customxmlpartcollection)|[add(xml: string)](/javascript/api/word/word.customxmlpartcollection#add-xml-)|Adds a new custom XML part to the document.|
||[getByNamespace(namespaceUri: string)](/javascript/api/word/word.customxmlpartcollection#getbynamespace-namespaceuri-)|Gets a new scoped collection of custom XML parts whose namespaces match the given namespace.|
||[getCount()](/javascript/api/word/word.customxmlpartcollection#getcount--)|Gets the number of items in the collection.|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartcollection#getitem-id-)|Gets a custom XML part based on its ID. Read only.|
||[getItemOrNullObject(id: string)](/javascript/api/word/word.customxmlpartcollection#getitemornullobject-id-)|Gets a custom XML part based on its ID. Returns a null object if the CustomXmlPart does not exist.|
||[items](/javascript/api/word/word.customxmlpartcollection#items)|Gets the loaded child items in this collection.|
|[CustomXmlPartScopedCollection](/javascript/api/word/word.customxmlpartscopedcollection)|[getCount()](/javascript/api/word/word.customxmlpartscopedcollection#getcount--)|Gets the number of items in the collection.|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartscopedcollection#getitem-id-)|Gets a custom XML part based on its ID. Read only.|
||[getItemOrNullObject(id: string)](/javascript/api/word/word.customxmlpartscopedcollection#getitemornullobject-id-)|Gets a custom XML part based on its ID. Returns a null object if the CustomXmlPart does not exist in the collection.|
||[getOnlyItem()](/javascript/api/word/word.customxmlpartscopedcollection#getonlyitem--)|If the collection contains exactly one item, this method returns it. Otherwise, this method produces an error.|
||[getOnlyItemOrNullObject()](/javascript/api/word/word.customxmlpartscopedcollection#getonlyitemornullobject--)|If the collection contains exactly one item, this method returns it. Otherwise, this method returns a null object.|
||[items](/javascript/api/word/word.customxmlpartscopedcollection#items)|Gets the loaded child items in this collection.|
|[Document](/javascript/api/word/word.document)|[deleteBookmark(name: string)](/javascript/api/word/word.document#deletebookmark-name-)|Deletes a bookmark, if exists, from the document.|
||[getBookmarkRange(name: string)](/javascript/api/word/word.document#getbookmarkrange-name-)|Gets a bookmark's range. Throws if the bookmark does not exist.|
||[getBookmarkRangeOrNullObject(name: string)](/javascript/api/word/word.document#getbookmarkrangeornullobject-name-)|Gets a bookmark's range. Returns a null object if the bookmark does not exist.|
||[customXmlParts](/javascript/api/word/word.document#customxmlparts)|Gets the custom XML parts in the document. Read-only.|
||[onContentControlAdded](/javascript/api/word/word.document#oncontentcontroladded)|Occurs when a content control is added. Run context.sync() in the handler to get the new content control's properties.|
||[settings](/javascript/api/word/word.document#settings)|Gets the add-in's settings in the document. Read-only.|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[deleteBookmark(name: string)](/javascript/api/word/word.documentcreated#deletebookmark-name-)|Deletes a bookmark, if exists, from the document.|
||[getBookmarkRange(name: string)](/javascript/api/word/word.documentcreated#getbookmarkrange-name-)|Gets a bookmark's range. Throws if the bookmark does not exist.|
||[getBookmarkRangeOrNullObject(name: string)](/javascript/api/word/word.documentcreated#getbookmarkrangeornullobject-name-)|Gets a bookmark's range. Returns a null object if the bookmark does not exist.|
||[customXmlParts](/javascript/api/word/word.documentcreated#customxmlparts)|Gets the custom XML parts in the document. Read-only.|
||[settings](/javascript/api/word/word.documentcreated#settings)|Gets the add-in's settings in the document. Read-only.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[imageFormat](/javascript/api/word/word.inlinepicture#imageformat)|Gets the format of the inline image. Read-only.|
|[List](/javascript/api/word/word.list)|[getLevelFont(level: number)](/javascript/api/word/word.list#getlevelfont-level-)|Gets the font of the bullet, number or picture at the specified level in the list.|
||[getLevelPicture(level: number)](/javascript/api/word/word.list#getlevelpicture-level-)|Gets the base64 encoded string representation of the picture at the specified level in the list.|
||[resetLevelFont(level: number, resetFontName?: boolean)](/javascript/api/word/word.list#resetlevelfont-level--resetfontname-)|Resets the font of the bullet, number or picture at the specified level in the list.|
||[setLevelPicture(level: number, base64EncodedImage?: string)](/javascript/api/word/word.list#setlevelpicture-level--base64encodedimage-)|Sets the picture at the specified level in the list.|
|[Range](/javascript/api/word/word.range)|[getBookmarks(includeHidden?: boolean, includeAdjacent?: boolean)](/javascript/api/word/word.range#getbookmarks-includehidden--includeadjacent-)|Gets the names all bookmarks in or overlapping the range. A bookmark is hidden if its name starts with the underscore character.|
||[insertBookmark(name: string)](/javascript/api/word/word.range#insertbookmark-name-)|Inserts a bookmark on the range. If a bookmark of the same name exists somewhere, it is deleted first.|
|[Setting](/javascript/api/word/word.setting)|[delete()](/javascript/api/word/word.setting#delete--)|Deletes the setting.|
||[key](/javascript/api/word/word.setting#key)|Gets the key of the setting. Read only.|
||[value](/javascript/api/word/word.setting#value)|Gets or sets the value of the setting.|
|[SettingCollection](/javascript/api/word/word.settingcollection)|[add(key: string, value: any)](/javascript/api/word/word.settingcollection#add-key--value-)|Creates a new setting or sets an existing setting.|
||[deleteAll()](/javascript/api/word/word.settingcollection#deleteall--)|Deletes all settings in this add-in.|
||[getCount()](/javascript/api/word/word.settingcollection#getcount--)|Gets the count of settings.|
||[getItem(key: string)](/javascript/api/word/word.settingcollection#getitem-key-)|Gets a setting object by its key, which is case-sensitive. Throws if the setting does not exist.|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.settingcollection#getitemornullobject-key-)|Gets a setting object by its key, which is case-sensitive. Returns a null object if the setting does not exist.|
||[items](/javascript/api/word/word.settingcollection#items)|Gets the loaded child items in this collection.|
|[Table](/javascript/api/word/word.table)|[mergeCells(topRow: number, firstCell: number, bottomRow: number, lastCell: number)](/javascript/api/word/word.table#mergecells-toprow--firstcell--bottomrow--lastcell-)|Merges the cells bounded inclusively by a first and last cell.|
|[TableCell](/javascript/api/word/word.tablecell)|[split(rowCount: number, columnCount: number)](/javascript/api/word/word.tablecell#split-rowcount--columncount-)|Splits the cell into the specified number of rows and columns.|
|[TableRow](/javascript/api/word/word.tablerow)|[insertContentControl()](/javascript/api/word/word.tablerow#insertcontentcontrol--)|Inserts a content control on the row.|
||[merge()](/javascript/api/word/word.tablerow#merge--)|Merges the row into one cell.|

## See also

- [Word JavaScript API Reference Documentation](/javascript/api/word)
- [Word JavaScript API requirement sets](word-api-requirement-sets.md)
