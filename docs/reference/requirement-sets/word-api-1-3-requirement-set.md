---
title: Word JavaScript API requirement set 1.3
description: 'Details about the WordApi 1.3 requirement set'
ms.date: 07/25/2019
ms.prod: word
localization_priority: Normal
---

# What's new in Word JavaScript API 1.3

WordApi 1.3 added more support for content controls, custom XML, and document-level settings.

## API list

The following table lists the APIs in Word JavaScript API requirement set 1.3. To view API reference documentation for all APIs supported by Word JavaScript API requirement set 1.3 or earlier, see [Word APIs in requirement set 1.3 or earlier](/javascript/api/word?view=word-js-1.3).

| Class | Fields | Description |
|:---|:---|:---|
|[Application](/javascript/api/word/word.application)|[createDocument(base64File?: string)](/javascript/api/word/word.application#createdocument-base64file-)|Creates a new document by using an optional base64 encoded .docx file.|
|[Body](/javascript/api/word/word.body)|[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.body#getrange-rangelocation-)|Gets the whole body, or the starting or ending point of the body, as a range.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.body#inserttable-rowcount--columncount--insertlocation--values-)|Inserts a table with the specified number of rows and columns. The insertLocation value can be 'Start' or 'End'.|
||[lists](/javascript/api/word/word.body#lists)|Gets the collection of list objects in the body. Read-only.|
||[parentBody](/javascript/api/word/word.body#parentbody)|Gets the parent body of the body. For example, a table cell body's parent body could be a header. Throws if there isn't a parent body. Read-only.|
||[parentBodyOrNullObject](/javascript/api/word/word.body#parentbodyornullobject)|Gets the parent body of the body. For example, a table cell body's parent body could be a header. Returns a null object if there isn't a parent body. Read-only.|
||[parentContentControlOrNullObject](/javascript/api/word/word.body#parentcontentcontrolornullobject)|Gets the content control that contains the body. Returns a null object if there isn't a parent content control. Read-only.|
||[parentSection](/javascript/api/word/word.body#parentsection)|Gets the parent section of the body. Throws if there isn't a parent section. Read-only.|
||[parentSectionOrNullObject](/javascript/api/word/word.body#parentsectionornullobject)|Gets the parent section of the body. Returns a null object if there isn't a parent section. Read-only.|
||[tables](/javascript/api/word/word.body#tables)|Gets the collection of table objects in the body. Read-only.|
||[type](/javascript/api/word/word.body#type)|Gets the type of the body. The type can be 'MainDoc', 'Section', 'Header', 'Footer', or 'TableCell'. Read-only.|
||[styleBuiltIn](/javascript/api/word/word.body#stylebuiltin)|Gets or sets the built-in style name for the body. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.contentcontrol#getrange-rangelocation-)|Gets the whole content control, or the starting or ending point of the content control, as a range.|
||[getTextRanges(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.contentcontrol#gettextranges-endingmarks--trimspacing-)|Gets the text ranges in the content control by using punctuation marks and/or other ending marks.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.contentcontrol#inserttable-rowcount--columncount--insertlocation--values-)|Inserts a table with the specified number of rows and columns into, or next to, a content control. The insertLocation value can be 'Start', 'End', 'Before', or 'After'.|
||[lists](/javascript/api/word/word.contentcontrol#lists)|Gets the collection of list objects in the content control. Read-only.|
||[parentBody](/javascript/api/word/word.contentcontrol#parentbody)|Gets the parent body of the content control. Read-only.|
||[parentContentControlOrNullObject](/javascript/api/word/word.contentcontrol#parentcontentcontrolornullobject)|Gets the content control that contains the content control. Returns a null object if there isn't a parent content control. Read-only.|
||[parentTable](/javascript/api/word/word.contentcontrol#parenttable)|Gets the table that contains the content control. Throws if it is not contained in a table. Read-only.|
||[parentTableCell](/javascript/api/word/word.contentcontrol#parenttablecell)|Gets the table cell that contains the content control. Throws if it is not contained in a table cell. Read-only.|
||[parentTableCellOrNullObject](/javascript/api/word/word.contentcontrol#parenttablecellornullobject)|Gets the table cell that contains the content control. Returns a null object if it is not contained in a table cell. Read-only.|
||[parentTableOrNullObject](/javascript/api/word/word.contentcontrol#parenttableornullobject)|Gets the table that contains the content control. Returns a null object if it is not contained in a table. Read-only.|
||[subtype](/javascript/api/word/word.contentcontrol#subtype)|Gets the content control subtype. The subtype can be 'RichTextInline', 'RichTextParagraphs', 'RichTextTableCell', 'RichTextTableRow' and 'RichTextTable' for rich text content controls. Read-only.|
||[tables](/javascript/api/word/word.contentcontrol#tables)|Gets the collection of table objects in the content control. Read-only.|
||[split(delimiters: string[], multiParagraphs?: boolean, trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.contentcontrol#split-delimiters--multiparagraphs--trimdelimiters--trimspacing-)|Splits the content control into child ranges by using delimiters.|
||[styleBuiltIn](/javascript/api/word/word.contentcontrol#stylebuiltin)|Gets or sets the built-in style name for the content control. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getByIdOrNullObject(id: number)](/javascript/api/word/word.contentcontrolcollection#getbyidornullobject-id-)|Gets a content control by its identifier. Returns a null object if there isn't a content control with the identifier in this collection.|
||[getByTypes(types: Word.ContentControlType[])](/javascript/api/word/word.contentcontrolcollection#getbytypes-types-)|Gets the content controls that have the specified types and/or subtypes.|
||[getFirst()](/javascript/api/word/word.contentcontrolcollection#getfirst--)|Gets the first content control in this collection. Throws if this collection is empty.|
||[getFirstOrNullObject()](/javascript/api/word/word.contentcontrolcollection#getfirstornullobject--)|Gets the first content control in this collection. Returns a null object if this collection is empty.|
|[CustomProperty](/javascript/api/word/word.customproperty)|[delete()](/javascript/api/word/word.customproperty#delete--)|Deletes the custom property.|
||[key](/javascript/api/word/word.customproperty#key)|Gets the key of the custom property. Read only.|
||[type](/javascript/api/word/word.customproperty#type)|Gets the value type of the custom property. Possible values are: String, Number, Date, Boolean. Read only.|
||[value](/javascript/api/word/word.customproperty#value)|Gets or sets the value of the custom property. Note that even though Word on the web and the docx file format allow these properties to be arbitrarily long, the desktop version of Word will truncate string values to 255 16-bit chars (possibly creating invalid unicode by breaking up a surrogate pair).|
|[CustomPropertyCollection](/javascript/api/word/word.custompropertycollection)|[add(key: string, value: any)](/javascript/api/word/word.custompropertycollection#add-key--value-)|Creates a new or sets an existing custom property.|
||[deleteAll()](/javascript/api/word/word.custompropertycollection#deleteall--)|Deletes all custom properties in this collection.|
||[getCount()](/javascript/api/word/word.custompropertycollection#getcount--)|Gets the count of custom properties.|
||[getItem(key: string)](/javascript/api/word/word.custompropertycollection#getitem-key-)|Gets a custom property object by its key, which is case-insensitive. Throws if the custom property does not exist.|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.custompropertycollection#getitemornullobject-key-)|Gets a custom property object by its key, which is case-insensitive. Returns a null object if the custom property does not exist.|
||[items](/javascript/api/word/word.custompropertycollection#items)|Gets the loaded child items in this collection.|
|[Document](/javascript/api/word/word.document)|[properties](/javascript/api/word/word.document#properties)|Gets the properties of the document. Read-only.|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[open()](/javascript/api/word/word.documentcreated#open--)|Opens the document.|
||[body](/javascript/api/word/word.documentcreated#body)|Gets the body object of the document. The body is the text that excludes headers, footers, footnotes, textboxes, etc.. Read-only.|
||[contentControls](/javascript/api/word/word.documentcreated#contentcontrols)|Gets the collection of content control objects in the document. This includes content controls in the body of the document, headers, footers, textboxes, etc.. Read-only.|
||[properties](/javascript/api/word/word.documentcreated#properties)|Gets the properties of the document. Read-only.|
||[saved](/javascript/api/word/word.documentcreated#saved)|Indicates whether the changes in the document have been saved. A value of true indicates that the document hasn't changed since it was saved. Read-only.|
||[sections](/javascript/api/word/word.documentcreated#sections)|Gets the collection of section objects in the document. Read-only.|
||[save()](/javascript/api/word/word.documentcreated#save--)|Saves the document. This will use the Word default file naming convention if the document has not been saved before.|
|[DocumentProperties](/javascript/api/word/word.documentproperties)|[author](/javascript/api/word/word.documentproperties#author)|Gets or sets the author of the document.|
||[category](/javascript/api/word/word.documentproperties#category)|Gets or sets the category of the document.|
||[comments](/javascript/api/word/word.documentproperties#comments)|Gets or sets the comments of the document.|
||[company](/javascript/api/word/word.documentproperties#company)|Gets or sets the company of the document.|
||[format](/javascript/api/word/word.documentproperties#format)|Gets or sets the format of the document.|
||[keywords](/javascript/api/word/word.documentproperties#keywords)|Gets or sets the keywords of the document.|
||[manager](/javascript/api/word/word.documentproperties#manager)|Gets or sets the manager of the document.|
||[applicationName](/javascript/api/word/word.documentproperties#applicationname)|Gets the application name of the document. Read only.|
||[creationDate](/javascript/api/word/word.documentproperties#creationdate)|Gets the creation date of the document. Read only.|
||[customProperties](/javascript/api/word/word.documentproperties#customproperties)|Gets the collection of custom properties of the document. Read only.|
||[lastAuthor](/javascript/api/word/word.documentproperties#lastauthor)|Gets the last author of the document. Read only.|
||[lastPrintDate](/javascript/api/word/word.documentproperties#lastprintdate)|Gets the last print date of the document. Read only.|
||[lastSaveTime](/javascript/api/word/word.documentproperties#lastsavetime)|Gets the last save time of the document. Read only.|
||[revisionNumber](/javascript/api/word/word.documentproperties#revisionnumber)|Gets the revision number of the document. Read only.|
||[security](/javascript/api/word/word.documentproperties#security)|Gets the security of the document. Read only.|
||[template](/javascript/api/word/word.documentproperties#template)|Gets the template of the document. Read only.|
||[subject](/javascript/api/word/word.documentproperties#subject)|Gets or sets the subject of the document.|
||[title](/javascript/api/word/word.documentproperties#title)|Gets or sets the title of the document.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[getNext()](/javascript/api/word/word.inlinepicture#getnext--)|Gets the next inline image. Throws if this inline image is the last one.|
||[getNextOrNullObject()](/javascript/api/word/word.inlinepicture#getnextornullobject--)|Gets the next inline image. Returns a null object if this inline image is the last one.|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.inlinepicture#getrange-rangelocation-)|Gets the picture, or the starting or ending point of the picture, as a range.|
||[parentContentControlOrNullObject](/javascript/api/word/word.inlinepicture#parentcontentcontrolornullobject)|Gets the content control that contains the inline image. Returns a null object if there isn't a parent content control. Read-only.|
||[parentTable](/javascript/api/word/word.inlinepicture#parenttable)|Gets the table that contains the inline image. Throws if it is not contained in a table. Read-only.|
||[parentTableCell](/javascript/api/word/word.inlinepicture#parenttablecell)|Gets the table cell that contains the inline image. Throws if it is not contained in a table cell. Read-only.|
||[parentTableCellOrNullObject](/javascript/api/word/word.inlinepicture#parenttablecellornullobject)|Gets the table cell that contains the inline image. Returns a null object if it is not contained in a table cell. Read-only.|
||[parentTableOrNullObject](/javascript/api/word/word.inlinepicture#parenttableornullobject)|Gets the table that contains the inline image. Returns a null object if it is not contained in a table. Read-only.|
|[InlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|[getFirst()](/javascript/api/word/word.inlinepicturecollection#getfirst--)|Gets the first inline image in this collection. Throws if this collection is empty.|
||[getFirstOrNullObject()](/javascript/api/word/word.inlinepicturecollection#getfirstornullobject--)|Gets the first inline image in this collection. Returns a null object if this collection is empty.|
|[List](/javascript/api/word/word.list)|[getLevelParagraphs(level: number)](/javascript/api/word/word.list#getlevelparagraphs-level-)|Gets the paragraphs that occur at the specified level in the list.|
||[getLevelString(level: number)](/javascript/api/word/word.list#getlevelstring-level-)|Gets the bullet, number or picture at the specified level as a string.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.list#insertparagraph-paragraphtext--insertlocation-)|Inserts a paragraph at the specified location. The insertLocation value can be 'Start', 'End', 'Before', or 'After'.|
||[id](/javascript/api/word/word.list#id)|Gets the list's id.|
||[levelExistences](/javascript/api/word/word.list#levelexistences)|Checks whether each of the 9 levels exists in the list. A true value indicates the level exists, which means there is at least one list item at that level. Read-only.|
||[levelTypes](/javascript/api/word/word.list#leveltypes)|Gets all 9 level types in the list. Each type can be 'Bullet', 'Number', or 'Picture'. Read-only.|
||[paragraphs](/javascript/api/word/word.list#paragraphs)|Gets paragraphs in the list. Read-only.|
||[setLevelAlignment(level: number, alignment: Word.Alignment)](/javascript/api/word/word.list#setlevelalignment-level--alignment-)|Sets the alignment of the bullet, number or picture at the specified level in the list.|
||[setLevelBullet(level: number, listBullet: Word.ListBullet, charCode?: number, fontName?: string)](/javascript/api/word/word.list#setlevelbullet-level--listbullet--charcode--fontname-)|Sets the bullet format at the specified level in the list. If the bullet is 'Custom', the charCode is required.|
||[setLevelIndents(level: number, textIndent: number, bulletNumberPictureIndent: number)](/javascript/api/word/word.list#setlevelindents-level--textindent--bulletnumberpictureindent-)|Sets the two indents of the specified level in the list.|
||[setLevelNumbering(level: number, listNumbering: Word.ListNumbering, formatString?: Array<string \| number>)](/javascript/api/word/word.list#setlevelnumbering-level--listnumbering--formatstring-)|Sets the numbering format at the specified level in the list.|
||[setLevelStartingNumber(level: number, startingNumber: number)](/javascript/api/word/word.list#setlevelstartingnumber-level--startingnumber-)|Sets the starting number at the specified level in the list. Default value is 1.|
|[ListCollection](/javascript/api/word/word.listcollection)|[getById(id: number)](/javascript/api/word/word.listcollection#getbyid-id-)|Gets a list by its identifier. Throws if there isn't a list with the identifier in this collection.|
||[getByIdOrNullObject(id: number)](/javascript/api/word/word.listcollection#getbyidornullobject-id-)|Gets a list by its identifier. Returns a null object if there isn't a list with the identifier in this collection.|
||[getFirst()](/javascript/api/word/word.listcollection#getfirst--)|Gets the first list in this collection. Throws if this collection is empty.|
||[getFirstOrNullObject()](/javascript/api/word/word.listcollection#getfirstornullobject--)|Gets the first list in this collection. Returns a null object if this collection is empty.|
||[getItem(index: number)](/javascript/api/word/word.listcollection#getitem-index-)|Gets a list object by its index in the collection.|
||[items](/javascript/api/word/word.listcollection#items)|Gets the loaded child items in this collection.|
|[ListItem](/javascript/api/word/word.listitem)|[getAncestor(parentOnly?: boolean)](/javascript/api/word/word.listitem#getancestor-parentonly-)|Gets the list item parent, or the closest ancestor if the parent does not exist. Throws if the list item has no ancestor.|
||[getAncestorOrNullObject(parentOnly?: boolean)](/javascript/api/word/word.listitem#getancestorornullobject-parentonly-)|Gets the list item parent, or the closest ancestor if the parent does not exist. Returns a null object if the list item has no ancestor.|
||[getDescendants(directChildrenOnly?: boolean)](/javascript/api/word/word.listitem#getdescendants-directchildrenonly-)|Gets all descendant list items of the list item.|
||[level](/javascript/api/word/word.listitem#level)|Gets or sets the level of the item in the list.|
||[listString](/javascript/api/word/word.listitem#liststring)|Gets the list item bullet, number, or picture as a string. Read-only.|
||[siblingIndex](/javascript/api/word/word.listitem#siblingindex)|Gets the list item order number in relation to its siblings. Read-only.|
|[Paragraph](/javascript/api/word/word.paragraph)|[attachToList(listId: number, level: number)](/javascript/api/word/word.paragraph#attachtolist-listid--level-)|Lets the paragraph join an existing list at the specified level. Fails if the paragraph cannot join the list or if the paragraph is already a list item.|
||[detachFromList()](/javascript/api/word/word.paragraph#detachfromlist--)|Moves this paragraph out of its list, if the paragraph is a list item.|
||[getNext()](/javascript/api/word/word.paragraph#getnext--)|Gets the next paragraph. Throws if the paragraph is the last one.|
||[getNextOrNullObject()](/javascript/api/word/word.paragraph#getnextornullobject--)|Gets the next paragraph. Returns a null object if the paragraph is the last one.|
||[getPrevious()](/javascript/api/word/word.paragraph#getprevious--)|Gets the previous paragraph. Throws if the paragraph is the first one.|
||[getPreviousOrNullObject()](/javascript/api/word/word.paragraph#getpreviousornullobject--)|Gets the previous paragraph. Returns a null object if the paragraph is the first one.|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.paragraph#getrange-rangelocation-)|Gets the whole paragraph, or the starting or ending point of the paragraph, as a range.|
||[getTextRanges(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.paragraph#gettextranges-endingmarks--trimspacing-)|Gets the text ranges in the paragraph by using punctuation marks and/or other ending marks.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.paragraph#inserttable-rowcount--columncount--insertlocation--values-)|Inserts a table with the specified number of rows and columns. The insertLocation value can be 'Before' or 'After'.|
||[isLastParagraph](/javascript/api/word/word.paragraph#islastparagraph)|Indicates the paragraph is the last one inside its parent body. Read-only.|
||[isListItem](/javascript/api/word/word.paragraph#islistitem)|Checks whether the paragraph is a list item. Read-only.|
||[list](/javascript/api/word/word.paragraph#list)|Gets the List to which this paragraph belongs. Throws if the paragraph is not in a list. Read-only.|
||[listItem](/javascript/api/word/word.paragraph#listitem)|Gets the ListItem for the paragraph. Throws if the paragraph is not part of a list. Read-only.|
||[listItemOrNullObject](/javascript/api/word/word.paragraph#listitemornullobject)|Gets the ListItem for the paragraph. Returns a null object if the paragraph is not part of a list. Read-only.|
||[listOrNullObject](/javascript/api/word/word.paragraph#listornullobject)|Gets the List to which this paragraph belongs. Returns a null object if the paragraph is not in a list. Read-only.|
||[parentBody](/javascript/api/word/word.paragraph#parentbody)|Gets the parent body of the paragraph. Read-only.|
||[parentContentControlOrNullObject](/javascript/api/word/word.paragraph#parentcontentcontrolornullobject)|Gets the content control that contains the paragraph. Returns a null object if there isn't a parent content control. Read-only.|
||[parentTable](/javascript/api/word/word.paragraph#parenttable)|Gets the table that contains the paragraph. Throws if it is not contained in a table. Read-only.|
||[parentTableCell](/javascript/api/word/word.paragraph#parenttablecell)|Gets the table cell that contains the paragraph. Throws if it is not contained in a table cell. Read-only.|
||[parentTableCellOrNullObject](/javascript/api/word/word.paragraph#parenttablecellornullobject)|Gets the table cell that contains the paragraph. Returns a null object if it is not contained in a table cell. Read-only.|
||[parentTableOrNullObject](/javascript/api/word/word.paragraph#parenttableornullobject)|Gets the table that contains the paragraph. Returns a null object if it is not contained in a table. Read-only.|
||[tableNestingLevel](/javascript/api/word/word.paragraph#tablenestinglevel)|Gets the level of the paragraph's table. It returns 0 if the paragraph is not in a table. Read-only.|
||[split(delimiters: string[], trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.paragraph#split-delimiters--trimdelimiters--trimspacing-)|Splits the paragraph into child ranges by using delimiters.|
||[startNewList()](/javascript/api/word/word.paragraph#startnewlist--)|Starts a new list with this paragraph. Fails if the paragraph is already a list item.|
||[styleBuiltIn](/javascript/api/word/word.paragraph#stylebuiltin)|Gets or sets the built-in style name for the paragraph. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[getFirst()](/javascript/api/word/word.paragraphcollection#getfirst--)|Gets the first paragraph in this collection. Throws if the collection is empty.|
||[getFirstOrNullObject()](/javascript/api/word/word.paragraphcollection#getfirstornullobject--)|Gets the first paragraph in this collection. Returns a null object if the collection is empty.|
||[getLast()](/javascript/api/word/word.paragraphcollection#getlast--)|Gets the last paragraph in this collection. Throws if the collection is empty.|
||[getLastOrNullObject()](/javascript/api/word/word.paragraphcollection#getlastornullobject--)|Gets the last paragraph in this collection. Returns a null object if the collection is empty.|
|[Range](/javascript/api/word/word.range)|[compareLocationWith(range: Word.Range)](/javascript/api/word/word.range#comparelocationwith-range-)|Compares this range's location with another range's location.|
||[expandTo(range: Word.Range)](/javascript/api/word/word.range#expandto-range-)|Returns a new range that extends from this range in either direction to cover another range. This range is not changed. Throws if the two ranges do not have a union.|
||[expandToOrNullObject(range: Word.Range)](/javascript/api/word/word.range#expandtoornullobject-range-)|Returns a new range that extends from this range in either direction to cover another range. This range is not changed. Returns a null object if the two ranges do not have a union.|
||[getHyperlinkRanges()](/javascript/api/word/word.range#gethyperlinkranges--)|Gets hyperlink child ranges within the range.|
||[getNextTextRange(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#getnexttextrange-endingmarks--trimspacing-)|Gets the next text range by using punctuation marks and/or other ending marks. Throws if this text range is the last one.|
||[getNextTextRangeOrNullObject(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#getnexttextrangeornullobject-endingmarks--trimspacing-)|Gets the next text range by using punctuation marks and/or other ending marks. Returns a null object if this text range is the last one.|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.range#getrange-rangelocation-)|Clones the range, or gets the starting or ending point of the range as a new range.|
||[getTextRanges(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#gettextranges-endingmarks--trimspacing-)|Gets the text child ranges in the range by using punctuation marks and/or other ending marks.|
||[hyperlink](/javascript/api/word/word.range#hyperlink)|Gets the first hyperlink in the range, or sets a hyperlink on the range. All hyperlinks in the range are deleted when you set a new hyperlink on the range. Use a '#' to separate the address part from the optional location part.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.range#inserttable-rowcount--columncount--insertlocation--values-)|Inserts a table with the specified number of rows and columns. The insertLocation value can be 'Before' or 'After'.|
||[intersectWith(range: Word.Range)](/javascript/api/word/word.range#intersectwith-range-)|Returns a new range as the intersection of this range with another range. This range is not changed. Throws if the two ranges are not overlapped or adjacent.|
||[intersectWithOrNullObject(range: Word.Range)](/javascript/api/word/word.range#intersectwithornullobject-range-)|Returns a new range as the intersection of this range with another range. This range is not changed. Returns a null object if the two ranges are not overlapped or adjacent.|
||[isEmpty](/javascript/api/word/word.range#isempty)|Checks whether the range length is zero. Read-only.|
||[lists](/javascript/api/word/word.range#lists)|Gets the collection of list objects in the range. Read-only.|
||[parentBody](/javascript/api/word/word.range#parentbody)|Gets the parent body of the range. Read-only.|
||[parentContentControlOrNullObject](/javascript/api/word/word.range#parentcontentcontrolornullobject)|Gets the content control that contains the range. Returns a null object if there isn't a parent content control. Read-only.|
||[parentTable](/javascript/api/word/word.range#parenttable)|Gets the table that contains the range. Throws if it is not contained in a table. Read-only.|
||[parentTableCell](/javascript/api/word/word.range#parenttablecell)|Gets the table cell that contains the range. Throws if it is not contained in a table cell. Read-only.|
||[parentTableCellOrNullObject](/javascript/api/word/word.range#parenttablecellornullobject)|Gets the table cell that contains the range. Returns a null object if it is not contained in a table cell. Read-only.|
||[parentTableOrNullObject](/javascript/api/word/word.range#parenttableornullobject)|Gets the table that contains the range. Returns a null object if it is not contained in a table. Read-only.|
||[tables](/javascript/api/word/word.range#tables)|Gets the collection of table objects in the range. Read-only.|
||[split(delimiters: string[], multiParagraphs?: boolean, trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.range#split-delimiters--multiparagraphs--trimdelimiters--trimspacing-)|Splits the range into child ranges by using delimiters.|
||[styleBuiltIn](/javascript/api/word/word.range#stylebuiltin)|Gets or sets the built-in style name for the range. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[getFirst()](/javascript/api/word/word.rangecollection#getfirst--)|Gets the first range in this collection. Throws if this collection is empty.|
||[getFirstOrNullObject()](/javascript/api/word/word.rangecollection#getfirstornullobject--)|Gets the first range in this collection. Returns a null object if this collection is empty.|
|[Section](/javascript/api/word/word.section)|[getNext()](/javascript/api/word/word.section#getnext--)|Gets the next section. Throws if this section is the last one.|
||[getNextOrNullObject()](/javascript/api/word/word.section#getnextornullobject--)|Gets the next section. Returns a null object if this section is the last one.|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[getFirst()](/javascript/api/word/word.sectioncollection#getfirst--)|Gets the first section in this collection. Throws if this collection is empty.|
||[getFirstOrNullObject()](/javascript/api/word/word.sectioncollection#getfirstornullobject--)|Gets the first section in this collection. Returns a null object if this collection is empty.|
|[Table](/javascript/api/word/word.table)|[addColumns(insertLocation: Word.InsertLocation, columnCount: number, values?: string[][])](/javascript/api/word/word.table#addcolumns-insertlocation--columncount--values-)|Adds columns to the start or end of the table, using the first or last existing column as a template. This is applicable to uniform tables. The string values, if specified, are set in the newly inserted rows.|
||[addRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[][])](/javascript/api/word/word.table#addrows-insertlocation--rowcount--values-)|Adds rows to the start or end of the table, using the first or last existing row as a template. The string values, if specified, are set in the newly inserted rows.|
||[alignment](/javascript/api/word/word.table#alignment)|Gets or sets the alignment of the table against the page column. The value can be 'Left', 'Centered', or 'Right'.|
||[autoFitWindow()](/javascript/api/word/word.table#autofitwindow--)|Autofits the table columns to the width of the window.|
||[clear()](/javascript/api/word/word.table#clear--)|Clears the contents of the table.|
||[delete()](/javascript/api/word/word.table#delete--)|Deletes the entire table.|
||[deleteColumns(columnIndex: number, columnCount?: number)](/javascript/api/word/word.table#deletecolumns-columnindex--columncount-)|Deletes specific columns. This is applicable to uniform tables.|
||[deleteRows(rowIndex: number, rowCount?: number)](/javascript/api/word/word.table#deleterows-rowindex--rowcount-)|Deletes specific rows.|
||[distributeColumns()](/javascript/api/word/word.table#distributecolumns--)|Distributes the column widths evenly. This is applicable to uniform tables.|
||[getBorder(borderLocation: Word.BorderLocation)](/javascript/api/word/word.table#getborder-borderlocation-)|Gets the border style for the specified border.|
||[getCell(rowIndex: number, cellIndex: number)](/javascript/api/word/word.table#getcell-rowindex--cellindex-)|Gets the table cell at a specified row and column. Throws if the specified table cell does not exist.|
||[getCellOrNullObject(rowIndex: number, cellIndex: number)](/javascript/api/word/word.table#getcellornullobject-rowindex--cellindex-)|Gets the table cell at a specified row and column. Returns a null object if the specified table cell does not exist.|
||[getCellPadding(cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.table#getcellpadding-cellpaddinglocation-)|Gets cell padding in points.|
||[getNext()](/javascript/api/word/word.table#getnext--)|Gets the next table. Throws if this table is the last one.|
||[getNextOrNullObject()](/javascript/api/word/word.table#getnextornullobject--)|Gets the next table. Returns a null object if this table is the last one.|
||[getParagraphAfter()](/javascript/api/word/word.table#getparagraphafter--)|Gets the paragraph after the table. Throws if there isn't a paragraph after the table.|
||[getParagraphAfterOrNullObject()](/javascript/api/word/word.table#getparagraphafterornullobject--)|Gets the paragraph after the table. Returns a null object if there isn't a paragraph after the table.|
||[getParagraphBefore()](/javascript/api/word/word.table#getparagraphbefore--)|Gets the paragraph before the table. Throws if there isn't a paragraph before the table.|
||[getParagraphBeforeOrNullObject()](/javascript/api/word/word.table#getparagraphbeforeornullobject--)|Gets the paragraph before the table. Returns a null object if there isn't a paragraph before the table.|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.table#getrange-rangelocation-)|Gets the range that contains this table, or the range at the start or end of the table.|
||[headerRowCount](/javascript/api/word/word.table#headerrowcount)|Gets and sets the number of header rows.|
||[horizontalAlignment](/javascript/api/word/word.table#horizontalalignment)|Gets and sets the horizontal alignment of every cell in the table. The value can be 'Left', 'Centered', 'Right', or 'Justified'.|
||[insertContentControl()](/javascript/api/word/word.table#insertcontentcontrol--)|Inserts a content control on the table.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.table#insertparagraph-paragraphtext--insertlocation-)|Inserts a paragraph at the specified location. The insertLocation value can be 'Before' or 'After'.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.table#inserttable-rowcount--columncount--insertlocation--values-)|Inserts a table with the specified number of rows and columns. The insertLocation value can be 'Before' or 'After'.|
||[font](/javascript/api/word/word.table#font)|Gets the font. Use this to get and set font name, size, color, and other properties. Read-only.|
||[isUniform](/javascript/api/word/word.table#isuniform)|Indicates whether all of the table rows are uniform. Read-only.|
||[nestingLevel](/javascript/api/word/word.table#nestinglevel)|Gets the nesting level of the table. Top-level tables have level 1. Read-only.|
||[parentBody](/javascript/api/word/word.table#parentbody)|Gets the parent body of the table. Read-only.|
||[parentContentControl](/javascript/api/word/word.table#parentcontentcontrol)|Gets the content control that contains the table. Throws if there isn't a parent content control. Read-only.|
||[parentContentControlOrNullObject](/javascript/api/word/word.table#parentcontentcontrolornullobject)|Gets the content control that contains the table. Returns a null object if there isn't a parent content control. Read-only.|
||[parentTable](/javascript/api/word/word.table#parenttable)|Gets the table that contains this table. Throws if it is not contained in a table. Read-only.|
||[parentTableCell](/javascript/api/word/word.table#parenttablecell)|Gets the table cell that contains this table. Throws if it is not contained in a table cell. Read-only.|
||[parentTableCellOrNullObject](/javascript/api/word/word.table#parenttablecellornullobject)|Gets the table cell that contains this table. Returns a null object if it is not contained in a table cell. Read-only.|
||[parentTableOrNullObject](/javascript/api/word/word.table#parenttableornullobject)|Gets the table that contains this table. Returns a null object if it is not contained in a table. Read-only.|
||[rowCount](/javascript/api/word/word.table#rowcount)|Gets the number of rows in the table. Read-only.|
||[rows](/javascript/api/word/word.table#rows)|Gets all of the table rows. Read-only.|
||[tables](/javascript/api/word/word.table#tables)|Gets the child tables nested one level deeper. Read-only.|
||[search(searchText: string, searchOptions?: Word.SearchOptions](/javascript/api/word/word.table#search-searchtext--searchoptions-)|Performs a search with the specified SearchOptions on the scope of the table object. The search results are a collection of range objects.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.table#select-selectionmode-)|Selects the table, or the position at the start or end of the table, and navigates the Word UI to it.|
||[setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.table#setcellpadding-cellpaddinglocation--cellpadding-)|Sets cell padding in points.|
||[shadingColor](/javascript/api/word/word.table#shadingcolor)|Gets and sets the shading color. Color is specified in "#RRGGBB" format or by using the color name.|
||[style](/javascript/api/word/word.table#style)|Gets or sets the style name for the table. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.|
||[styleBandedColumns](/javascript/api/word/word.table#stylebandedcolumns)|Gets and sets whether the table has banded columns.|
||[styleBandedRows](/javascript/api/word/word.table#stylebandedrows)|Gets and sets whether the table has banded rows.|
||[styleBuiltIn](/javascript/api/word/word.table#stylebuiltin)|Gets or sets the built-in style name for the table. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.|
||[styleFirstColumn](/javascript/api/word/word.table#stylefirstcolumn)|Gets and sets whether the table has a first column with a special style.|
||[styleLastColumn](/javascript/api/word/word.table#stylelastcolumn)|Gets and sets whether the table has a last column with a special style.|
||[styleTotalRow](/javascript/api/word/word.table#styletotalrow)|Gets and sets whether the table has a total (last) row with a special style.|
||[values](/javascript/api/word/word.table#values)|Gets and sets the text values in the table, as a 2D Javascript array.|
||[verticalAlignment](/javascript/api/word/word.table#verticalalignment)|Gets and sets the vertical alignment of every cell in the table. The value can be 'Top', 'Center', or 'Bottom'.|
||[width](/javascript/api/word/word.table#width)|Gets and sets the width of the table in points.|
|[TableBorder](/javascript/api/word/word.tableborder)|[color](/javascript/api/word/word.tableborder#color)|Gets or sets the table border color.|
||[type](/javascript/api/word/word.tableborder#type)|Gets or sets the type of the table border.|
||[width](/javascript/api/word/word.tableborder#width)|Gets or sets the width, in points, of the table border. Not applicable to table border types that have fixed widths.|
|[TableCell](/javascript/api/word/word.tablecell)|[columnWidth](/javascript/api/word/word.tablecell#columnwidth)|Gets and sets the width of the cell's column in points. This is applicable to uniform tables.|
||[deleteColumn()](/javascript/api/word/word.tablecell#deletecolumn--)|Deletes the column containing this cell. This is applicable to uniform tables.|
||[deleteRow()](/javascript/api/word/word.tablecell#deleterow--)|Deletes the row containing this cell.|
||[getBorder(borderLocation: Word.BorderLocation)](/javascript/api/word/word.tablecell#getborder-borderlocation-)|Gets the border style for the specified border.|
||[getCellPadding(cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.tablecell#getcellpadding-cellpaddinglocation-)|Gets cell padding in points.|
||[getNext()](/javascript/api/word/word.tablecell#getnext--)|Gets the next cell. Throws if this cell is the last one.|
||[getNextOrNullObject()](/javascript/api/word/word.tablecell#getnextornullobject--)|Gets the next cell. Returns a null object if this cell is the last one.|
||[horizontalAlignment](/javascript/api/word/word.tablecell#horizontalalignment)|Gets and sets the horizontal alignment of the cell. The value can be 'Left', 'Centered', 'Right', or 'Justified'.|
||[insertColumns(insertLocation: Word.InsertLocation, columnCount: number, values?: string[][])](/javascript/api/word/word.tablecell#insertcolumns-insertlocation--columncount--values-)|Adds columns to the left or right of the cell, using the cell's column as a template. This is applicable to uniform tables. The string values, if specified, are set in the newly inserted rows.|
||[insertRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[][])](/javascript/api/word/word.tablecell#insertrows-insertlocation--rowcount--values-)|Inserts rows above or below the cell, using the cell's row as a template. The string values, if specified, are set in the newly inserted rows.|
||[body](/javascript/api/word/word.tablecell#body)|Gets the body object of the cell. Read-only.|
||[cellIndex](/javascript/api/word/word.tablecell#cellindex)|Gets the index of the cell in its row. Read-only.|
||[parentRow](/javascript/api/word/word.tablecell#parentrow)|Gets the parent row of the cell. Read-only.|
||[parentTable](/javascript/api/word/word.tablecell#parenttable)|Gets the parent table of the cell. Read-only.|
||[rowIndex](/javascript/api/word/word.tablecell#rowindex)|Gets the index of the cell's row in the table. Read-only.|
||[width](/javascript/api/word/word.tablecell#width)|Gets the width of the cell in points. Read-only.|
||[setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.tablecell#setcellpadding-cellpaddinglocation--cellpadding-)|Sets cell padding in points.|
||[shadingColor](/javascript/api/word/word.tablecell#shadingcolor)|Gets or sets the shading color of the cell. Color is specified in "#RRGGBB" format or by using the color name.|
||[value](/javascript/api/word/word.tablecell#value)|Gets and sets the text of the cell.|
||[verticalAlignment](/javascript/api/word/word.tablecell#verticalalignment)|Gets and sets the vertical alignment of the cell. The value can be 'Top', 'Center', or 'Bottom'.|
|[TableCellCollection](/javascript/api/word/word.tablecellcollection)|[getFirst()](/javascript/api/word/word.tablecellcollection#getfirst--)|Gets the first table cell in this collection. Throws if this collection is empty.|
||[getFirstOrNullObject()](/javascript/api/word/word.tablecellcollection#getfirstornullobject--)|Gets the first table cell in this collection. Returns a null object if this collection is empty.|
||[items](/javascript/api/word/word.tablecellcollection#items)|Gets the loaded child items in this collection.|
|[TableCollection](/javascript/api/word/word.tablecollection)|[getFirst()](/javascript/api/word/word.tablecollection#getfirst--)|Gets the first table in this collection. Throws if this collection is empty.|
||[getFirstOrNullObject()](/javascript/api/word/word.tablecollection#getfirstornullobject--)|Gets the first table in this collection. Returns a null object if this collection is empty.|
||[items](/javascript/api/word/word.tablecollection#items)|Gets the loaded child items in this collection.|
|[TableRow](/javascript/api/word/word.tablerow)|[clear()](/javascript/api/word/word.tablerow#clear--)|Clears the contents of the row.|
||[delete()](/javascript/api/word/word.tablerow#delete--)|Deletes the entire row.|
||[getBorder(borderLocation: Word.BorderLocation)](/javascript/api/word/word.tablerow#getborder-borderlocation-)|Gets the border style of the cells in the row.|
||[getCellPadding(cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.tablerow#getcellpadding-cellpaddinglocation-)|Gets cell padding in points.|
||[getNext()](/javascript/api/word/word.tablerow#getnext--)|Gets the next row. Throws if this row is the last one.|
||[getNextOrNullObject()](/javascript/api/word/word.tablerow#getnextornullobject--)|Gets the next row. Returns a null object if this row is the last one.|
||[horizontalAlignment](/javascript/api/word/word.tablerow#horizontalalignment)|Gets and sets the horizontal alignment of every cell in the row. The value can be 'Left', 'Centered', 'Right', or 'Justified'.|
||[insertRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[][])](/javascript/api/word/word.tablerow#insertrows-insertlocation--rowcount--values-)|Inserts rows using this row as a template. If values are specified, inserts the values into the new rows.|
||[preferredHeight](/javascript/api/word/word.tablerow#preferredheight)|Gets and sets the preferred height of the row in points.|
||[cellCount](/javascript/api/word/word.tablerow#cellcount)|Gets the number of cells in the row. Read-only.|
||[cells](/javascript/api/word/word.tablerow#cells)|Gets cells. Read-only.|
||[font](/javascript/api/word/word.tablerow#font)|Gets the font. Use this to get and set font name, size, color, and other properties. Read-only.|
||[isHeader](/javascript/api/word/word.tablerow#isheader)|Checks whether the row is a header row. Read-only. To set the number of header rows, use HeaderRowCount on the Table object.|
||[parentTable](/javascript/api/word/word.tablerow#parenttable)|Gets parent table. Read-only.|
||[rowIndex](/javascript/api/word/word.tablerow#rowindex)|Gets the index of the row in its parent table. Read-only.|
||[search(searchText: string, searchOptions?: Word.SearchOptions)](/javascript/api/word/word.tablerow#search-searchtext--searchoptions-)|Performs a search with the specified SearchOptions on the scope of the row. The search results are a collection of range objects.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.tablerow#select-selectionmode-)|Selects the row and navigates the Word UI to it.|
||[setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.tablerow#setcellpadding-cellpaddinglocation--cellpadding-)|Sets cell padding in points.|
||[shadingColor](/javascript/api/word/word.tablerow#shadingcolor)|Gets and sets the shading color. Color is specified in "#RRGGBB" format or by using the color name.|
||[values](/javascript/api/word/word.tablerow#values)|Gets and sets the text values in the row, as a 2D Javascript array.|
||[verticalAlignment](/javascript/api/word/word.tablerow#verticalalignment)|Gets and sets the vertical alignment of the cells in the row. The value can be 'Top', 'Center', or 'Bottom'.|
|[TableRowCollection](/javascript/api/word/word.tablerowcollection)|[getFirst()](/javascript/api/word/word.tablerowcollection#getfirst--)|Gets the first row in this collection. Throws if this collection is empty.|
||[getFirstOrNullObject()](/javascript/api/word/word.tablerowcollection#getfirstornullobject--)|Gets the first row in this collection. Returns a null object if this collection is empty.|
||[items](/javascript/api/word/word.tablerowcollection#items)|Gets the loaded child items in this collection.|

## See also

- [Word JavaScript API Reference Documentation](/javascript/api/word)
- [Word JavaScript API requirement sets](word-api-requirement-sets.md)
