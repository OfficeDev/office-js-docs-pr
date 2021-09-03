---
title: Word JavaScript API requirement set 1.3
description: 'Details about the WordApi 1.3 requirement set.'
ms.date: 03/09/2021
ms.prod: word
localization_priority: Normal
---

# What's new in Word JavaScript API 1.3

WordApi 1.3 added more support for content controls and document-level settings.

## API list

The following table lists the APIs in Word JavaScript API requirement set 1.3. To view API reference documentation for all APIs supported by Word JavaScript API requirement set 1.3 or earlier, see [Word APIs in requirement set 1.3 or earlier](/javascript/api/word?view=word-js-1.3&preserve-view=true).

| Class | Fields | Description |
|:---|:---|:---|
|[Application](/javascript/api/word/word.application)|[createDocument(base64File?: string)](/javascript/api/word/word.application#createDocument_base64File_)|Creates a new document by using an optional base64 encoded .docx file.|
|[Body](/javascript/api/word/word.body)|[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.body#getRange_rangeLocation_)|Gets the whole body, or the starting or ending point of the body, as a range.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.body#insertTable_rowCount__columnCount__insertLocation__values_)|Inserts a table with the specified number of rows and columns.|
||[lists](/javascript/api/word/word.body#lists)|Gets the collection of list objects in the body.|
||[parentBody](/javascript/api/word/word.body#parentBody)|Gets the parent body of the body.|
||[parentBodyOrNullObject](/javascript/api/word/word.body#parentBodyOrNullObject)|Gets the parent body of the body.|
||[parentContentControlOrNullObject](/javascript/api/word/word.body#parentContentControlOrNullObject)|Gets the content control that contains the body.|
||[parentSection](/javascript/api/word/word.body#parentSection)|Gets the parent section of the body.|
||[parentSectionOrNullObject](/javascript/api/word/word.body#parentSectionOrNullObject)|Gets the parent section of the body.|
||[tables](/javascript/api/word/word.body#tables)|Gets the collection of table objects in the body.|
||[type](/javascript/api/word/word.body#type)|Gets the type of the body.|
||[styleBuiltIn](/javascript/api/word/word.body#styleBuiltIn)|Gets or sets the built-in style name for the body.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.contentcontrol#getRange_rangeLocation_)|Gets the whole content control, or the starting or ending point of the content control, as a range.|
||[getTextRanges(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.contentcontrol#getTextRanges_endingMarks__trimSpacing_)|Gets the text ranges in the content control by using punctuation marks and/or other ending marks.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.contentcontrol#insertTable_rowCount__columnCount__insertLocation__values_)|Inserts a table with the specified number of rows and columns into, or next to, a content control.|
||[lists](/javascript/api/word/word.contentcontrol#lists)|Gets the collection of list objects in the content control.|
||[parentBody](/javascript/api/word/word.contentcontrol#parentBody)|Gets the parent body of the content control.|
||[parentContentControlOrNullObject](/javascript/api/word/word.contentcontrol#parentContentControlOrNullObject)|Gets the content control that contains the content control.|
||[parentTable](/javascript/api/word/word.contentcontrol#parentTable)|Gets the table that contains the content control.|
||[parentTableCell](/javascript/api/word/word.contentcontrol#parentTableCell)|Gets the table cell that contains the content control.|
||[parentTableCellOrNullObject](/javascript/api/word/word.contentcontrol#parentTableCellOrNullObject)|Gets the table cell that contains the content control.|
||[parentTableOrNullObject](/javascript/api/word/word.contentcontrol#parentTableOrNullObject)|Gets the table that contains the content control.|
||[subtype](/javascript/api/word/word.contentcontrol#subtype)|Gets the content control subtype.|
||[tables](/javascript/api/word/word.contentcontrol#tables)|Gets the collection of table objects in the content control.|
||[split(delimiters: string[], multiParagraphs?: boolean, trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.contentcontrol#split_delimiters__multiParagraphs__trimDelimiters__trimSpacing_)|Splits the content control into child ranges by using delimiters.|
||[styleBuiltIn](/javascript/api/word/word.contentcontrol#styleBuiltIn)|Gets or sets the built-in style name for the content control.|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getByIdOrNullObject(id: number)](/javascript/api/word/word.contentcontrolcollection#getByIdOrNullObject_id_)|Gets a content control by its identifier.|
||[getByTypes(types: Word.ContentControlType[])](/javascript/api/word/word.contentcontrolcollection#getByTypes_types_)|Gets the content controls that have the specified types and/or subtypes.|
||[getFirst()](/javascript/api/word/word.contentcontrolcollection#getFirst__)|Gets the first content control in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.contentcontrolcollection#getFirstOrNullObject__)|Gets the first content control in this collection.|
|[CustomProperty](/javascript/api/word/word.customproperty)|[delete()](/javascript/api/word/word.customproperty#delete__)|Deletes the custom property.|
||[key](/javascript/api/word/word.customproperty#key)|Gets the key of the custom property.|
||[type](/javascript/api/word/word.customproperty#type)|Gets the value type of the custom property.|
||[value](/javascript/api/word/word.customproperty#value)|Gets or sets the value of the custom property.|
|[CustomPropertyCollection](/javascript/api/word/word.custompropertycollection)|[add(key: string, value: any)](/javascript/api/word/word.custompropertycollection#add_key__value_)|Creates a new or sets an existing custom property.|
||[deleteAll()](/javascript/api/word/word.custompropertycollection#deleteAll__)|Deletes all custom properties in this collection.|
||[getCount()](/javascript/api/word/word.custompropertycollection#getCount__)|Gets the count of custom properties.|
||[getItem(key: string)](/javascript/api/word/word.custompropertycollection#getItem_key_)|Gets a custom property object by its key, which is case-insensitive.|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.custompropertycollection#getItemOrNullObject_key_)|Gets a custom property object by its key, which is case-insensitive.|
||[items](/javascript/api/word/word.custompropertycollection#items)|Gets the loaded child items in this collection.|
|[Document](/javascript/api/word/word.document)|[properties](/javascript/api/word/word.document#properties)|Gets the properties of the document.|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[open()](/javascript/api/word/word.documentcreated#open__)|Opens the document.|
||[body](/javascript/api/word/word.documentcreated#body)|Gets the body object of the document.|
||[contentControls](/javascript/api/word/word.documentcreated#contentControls)|Gets the collection of content control objects in the document.|
||[properties](/javascript/api/word/word.documentcreated#properties)|Gets the properties of the document.|
||[saved](/javascript/api/word/word.documentcreated#saved)|Indicates whether the changes in the document have been saved.|
||[sections](/javascript/api/word/word.documentcreated#sections)|Gets the collection of section objects in the document.|
||[save()](/javascript/api/word/word.documentcreated#save__)|Saves the document.|
|[DocumentProperties](/javascript/api/word/word.documentproperties)|[author](/javascript/api/word/word.documentproperties#author)|Gets or sets the author of the document.|
||[category](/javascript/api/word/word.documentproperties#category)|Gets or sets the category of the document.|
||[comments](/javascript/api/word/word.documentproperties#comments)|Gets or sets the comments of the document.|
||[company](/javascript/api/word/word.documentproperties#company)|Gets or sets the company of the document.|
||[format](/javascript/api/word/word.documentproperties#format)|Gets or sets the format of the document.|
||[keywords](/javascript/api/word/word.documentproperties#keywords)|Gets or sets the keywords of the document.|
||[manager](/javascript/api/word/word.documentproperties#manager)|Gets or sets the manager of the document.|
||[applicationName](/javascript/api/word/word.documentproperties#applicationName)|Gets the application name of the document.|
||[creationDate](/javascript/api/word/word.documentproperties#creationDate)|Gets the creation date of the document.|
||[customProperties](/javascript/api/word/word.documentproperties#customProperties)|Gets the collection of custom properties of the document.|
||[lastAuthor](/javascript/api/word/word.documentproperties#lastAuthor)|Gets the last author of the document.|
||[lastPrintDate](/javascript/api/word/word.documentproperties#lastPrintDate)|Gets the last print date of the document.|
||[lastSaveTime](/javascript/api/word/word.documentproperties#lastSaveTime)|Gets the last save time of the document.|
||[revisionNumber](/javascript/api/word/word.documentproperties#revisionNumber)|Gets the revision number of the document.|
||[security](/javascript/api/word/word.documentproperties#security)|Gets security settings of the document.|
||[template](/javascript/api/word/word.documentproperties#template)|Gets the template of the document.|
||[subject](/javascript/api/word/word.documentproperties#subject)|Gets or sets the subject of the document.|
||[title](/javascript/api/word/word.documentproperties#title)|Gets or sets the title of the document.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[getNext()](/javascript/api/word/word.inlinepicture#getNext__)|Gets the next inline image.|
||[getNextOrNullObject()](/javascript/api/word/word.inlinepicture#getNextOrNullObject__)|Gets the next inline image.|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.inlinepicture#getRange_rangeLocation_)|Gets the picture, or the starting or ending point of the picture, as a range.|
||[parentContentControlOrNullObject](/javascript/api/word/word.inlinepicture#parentContentControlOrNullObject)|Gets the content control that contains the inline image.|
||[parentTable](/javascript/api/word/word.inlinepicture#parentTable)|Gets the table that contains the inline image.|
||[parentTableCell](/javascript/api/word/word.inlinepicture#parentTableCell)|Gets the table cell that contains the inline image.|
||[parentTableCellOrNullObject](/javascript/api/word/word.inlinepicture#parentTableCellOrNullObject)|Gets the table cell that contains the inline image.|
||[parentTableOrNullObject](/javascript/api/word/word.inlinepicture#parentTableOrNullObject)|Gets the table that contains the inline image.|
|[InlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|[getFirst()](/javascript/api/word/word.inlinepicturecollection#getFirst__)|Gets the first inline image in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.inlinepicturecollection#getFirstOrNullObject__)|Gets the first inline image in this collection.|
|[List](/javascript/api/word/word.list)|[getLevelParagraphs(level: number)](/javascript/api/word/word.list#getLevelParagraphs_level_)|Gets the paragraphs that occur at the specified level in the list.|
||[getLevelString(level: number)](/javascript/api/word/word.list#getLevelString_level_)|Gets the bullet, number, or picture at the specified level as a string.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.list#insertParagraph_paragraphText__insertLocation_)|Inserts a paragraph at the specified location.|
||[id](/javascript/api/word/word.list#id)|Gets the list's id.|
||[levelExistences](/javascript/api/word/word.list#levelExistences)|Checks whether each of the 9 levels exists in the list.|
||[levelTypes](/javascript/api/word/word.list#levelTypes)|Gets all 9 level types in the list.|
||[paragraphs](/javascript/api/word/word.list#paragraphs)|Gets paragraphs in the list.|
||[setLevelAlignment(level: number, alignment: Word.Alignment)](/javascript/api/word/word.list#setLevelAlignment_level__alignment_)|Sets the alignment of the bullet, number, or picture at the specified level in the list.|
||[setLevelBullet(level: number, listBullet: Word.ListBullet, charCode?: number, fontName?: string)](/javascript/api/word/word.list#setLevelBullet_level__listBullet__charCode__fontName_)|Sets the bullet format at the specified level in the list.|
||[setLevelIndents(level: number, textIndent: number, bulletNumberPictureIndent: number)](/javascript/api/word/word.list#setLevelIndents_level__textIndent__bulletNumberPictureIndent_)|Sets the two indents of the specified level in the list.|
||[setLevelNumbering(level: number, listNumbering: Word.ListNumbering, formatString?: Array<string \| number>)](/javascript/api/word/word.list#setLevelNumbering_level__listNumbering__formatString_)|Sets the numbering format at the specified level in the list.|
||[setLevelStartingNumber(level: number, startingNumber: number)](/javascript/api/word/word.list#setLevelStartingNumber_level__startingNumber_)|Sets the starting number at the specified level in the list.|
|[ListCollection](/javascript/api/word/word.listcollection)|[getById(id: number)](/javascript/api/word/word.listcollection#getById_id_)|Gets a list by its identifier.|
||[getByIdOrNullObject(id: number)](/javascript/api/word/word.listcollection#getByIdOrNullObject_id_)|Gets a list by its identifier.|
||[getFirst()](/javascript/api/word/word.listcollection#getFirst__)|Gets the first list in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.listcollection#getFirstOrNullObject__)|Gets the first list in this collection.|
||[getItem(index: number)](/javascript/api/word/word.listcollection#getItem_index_)|Gets a list object by its index in the collection.|
||[items](/javascript/api/word/word.listcollection#items)|Gets the loaded child items in this collection.|
|[ListItem](/javascript/api/word/word.listitem)|[getAncestor(parentOnly?: boolean)](/javascript/api/word/word.listitem#getAncestor_parentOnly_)|Gets the list item parent, or the closest ancestor if the parent does not exist.|
||[getAncestorOrNullObject(parentOnly?: boolean)](/javascript/api/word/word.listitem#getAncestorOrNullObject_parentOnly_)|Gets the list item parent, or the closest ancestor if the parent does not exist.|
||[getDescendants(directChildrenOnly?: boolean)](/javascript/api/word/word.listitem#getDescendants_directChildrenOnly_)|Gets all descendant list items of the list item.|
||[level](/javascript/api/word/word.listitem#level)|Gets or sets the level of the item in the list.|
||[listString](/javascript/api/word/word.listitem#listString)|Gets the list item bullet, number, or picture as a string.|
||[siblingIndex](/javascript/api/word/word.listitem#siblingIndex)|Gets the list item order number in relation to its siblings.|
|[Paragraph](/javascript/api/word/word.paragraph)|[attachToList(listId: number, level: number)](/javascript/api/word/word.paragraph#attachToList_listId__level_)|Lets the paragraph join an existing list at the specified level.|
||[detachFromList()](/javascript/api/word/word.paragraph#detachFromList__)|Moves this paragraph out of its list, if the paragraph is a list item.|
||[getNext()](/javascript/api/word/word.paragraph#getNext__)|Gets the next paragraph.|
||[getNextOrNullObject()](/javascript/api/word/word.paragraph#getNextOrNullObject__)|Gets the next paragraph.|
||[getPrevious()](/javascript/api/word/word.paragraph#getPrevious__)|Gets the previous paragraph.|
||[getPreviousOrNullObject()](/javascript/api/word/word.paragraph#getPreviousOrNullObject__)|Gets the previous paragraph.|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.paragraph#getRange_rangeLocation_)|Gets the whole paragraph, or the starting or ending point of the paragraph, as a range.|
||[getTextRanges(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.paragraph#getTextRanges_endingMarks__trimSpacing_)|Gets the text ranges in the paragraph by using punctuation marks and/or other ending marks.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.paragraph#insertTable_rowCount__columnCount__insertLocation__values_)|Inserts a table with the specified number of rows and columns.|
||[isLastParagraph](/javascript/api/word/word.paragraph#isLastParagraph)|Indicates the paragraph is the last one inside its parent body.|
||[isListItem](/javascript/api/word/word.paragraph#isListItem)|Checks whether the paragraph is a list item.|
||[list](/javascript/api/word/word.paragraph#list)|Gets the List to which this paragraph belongs.|
||[listItem](/javascript/api/word/word.paragraph#listItem)|Gets the ListItem for the paragraph.|
||[listItemOrNullObject](/javascript/api/word/word.paragraph#listItemOrNullObject)|Gets the ListItem for the paragraph.|
||[listOrNullObject](/javascript/api/word/word.paragraph#listOrNullObject)|Gets the List to which this paragraph belongs.|
||[parentBody](/javascript/api/word/word.paragraph#parentBody)|Gets the parent body of the paragraph.|
||[parentContentControlOrNullObject](/javascript/api/word/word.paragraph#parentContentControlOrNullObject)|Gets the content control that contains the paragraph.|
||[parentTable](/javascript/api/word/word.paragraph#parentTable)|Gets the table that contains the paragraph.|
||[parentTableCell](/javascript/api/word/word.paragraph#parentTableCell)|Gets the table cell that contains the paragraph.|
||[parentTableCellOrNullObject](/javascript/api/word/word.paragraph#parentTableCellOrNullObject)|Gets the table cell that contains the paragraph.|
||[parentTableOrNullObject](/javascript/api/word/word.paragraph#parentTableOrNullObject)|Gets the table that contains the paragraph.|
||[tableNestingLevel](/javascript/api/word/word.paragraph#tableNestingLevel)|Gets the level of the paragraph's table.|
||[split(delimiters: string[], trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.paragraph#split_delimiters__trimDelimiters__trimSpacing_)|Splits the paragraph into child ranges by using delimiters.|
||[startNewList()](/javascript/api/word/word.paragraph#startNewList__)|Starts a new list with this paragraph.|
||[styleBuiltIn](/javascript/api/word/word.paragraph#styleBuiltIn)|Gets or sets the built-in style name for the paragraph.|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[getFirst()](/javascript/api/word/word.paragraphcollection#getFirst__)|Gets the first paragraph in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.paragraphcollection#getFirstOrNullObject__)|Gets the first paragraph in this collection.|
||[getLast()](/javascript/api/word/word.paragraphcollection#getLast__)|Gets the last paragraph in this collection.|
||[getLastOrNullObject()](/javascript/api/word/word.paragraphcollection#getLastOrNullObject__)|Gets the last paragraph in this collection.|
|[Range](/javascript/api/word/word.range)|[compareLocationWith(range: Word.Range)](/javascript/api/word/word.range#compareLocationWith_range_)|Compares this range's location with another range's location.|
||[expandTo(range: Word.Range)](/javascript/api/word/word.range#expandTo_range_)|Returns a new range that extends from this range in either direction to cover another range.|
||[expandToOrNullObject(range: Word.Range)](/javascript/api/word/word.range#expandToOrNullObject_range_)|Returns a new range that extends from this range in either direction to cover another range.|
||[getHyperlinkRanges()](/javascript/api/word/word.range#getHyperlinkRanges__)|Gets hyperlink child ranges within the range.|
||[getNextTextRange(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#getNextTextRange_endingMarks__trimSpacing_)|Gets the next text range by using punctuation marks and/or other ending marks.|
||[getNextTextRangeOrNullObject(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#getNextTextRangeOrNullObject_endingMarks__trimSpacing_)|Gets the next text range by using punctuation marks and/or other ending marks.|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.range#getRange_rangeLocation_)|Clones the range, or gets the starting or ending point of the range as a new range.|
||[getTextRanges(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#getTextRanges_endingMarks__trimSpacing_)|Gets the text child ranges in the range by using punctuation marks and/or other ending marks.|
||[hyperlink](/javascript/api/word/word.range#hyperlink)|Gets the first hyperlink in the range, or sets a hyperlink on the range.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.range#insertTable_rowCount__columnCount__insertLocation__values_)|Inserts a table with the specified number of rows and columns.|
||[intersectWith(range: Word.Range)](/javascript/api/word/word.range#intersectWith_range_)|Returns a new range as the intersection of this range with another range.|
||[intersectWithOrNullObject(range: Word.Range)](/javascript/api/word/word.range#intersectWithOrNullObject_range_)|Returns a new range as the intersection of this range with another range.|
||[isEmpty](/javascript/api/word/word.range#isEmpty)|Checks whether the range length is zero.|
||[lists](/javascript/api/word/word.range#lists)|Gets the collection of list objects in the range.|
||[parentBody](/javascript/api/word/word.range#parentBody)|Gets the parent body of the range.|
||[parentContentControlOrNullObject](/javascript/api/word/word.range#parentContentControlOrNullObject)|Gets the content control that contains the range.|
||[parentTable](/javascript/api/word/word.range#parentTable)|Gets the table that contains the range.|
||[parentTableCell](/javascript/api/word/word.range#parentTableCell)|Gets the table cell that contains the range.|
||[parentTableCellOrNullObject](/javascript/api/word/word.range#parentTableCellOrNullObject)|Gets the table cell that contains the range.|
||[parentTableOrNullObject](/javascript/api/word/word.range#parentTableOrNullObject)|Gets the table that contains the range.|
||[tables](/javascript/api/word/word.range#tables)|Gets the collection of table objects in the range.|
||[split(delimiters: string[], multiParagraphs?: boolean, trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.range#split_delimiters__multiParagraphs__trimDelimiters__trimSpacing_)|Splits the range into child ranges by using delimiters.|
||[styleBuiltIn](/javascript/api/word/word.range#styleBuiltIn)|Gets or sets the built-in style name for the range.|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[getFirst()](/javascript/api/word/word.rangecollection#getFirst__)|Gets the first range in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.rangecollection#getFirstOrNullObject__)|Gets the first range in this collection.|
|[RequestContext](/javascript/api/word/word.requestcontext)|[application](/javascript/api/word/word.requestcontext#application)|[Api set: WordApi 1.3] *|
|[Section](/javascript/api/word/word.section)|[getNext()](/javascript/api/word/word.section#getNext__)|Gets the next section.|
||[getNextOrNullObject()](/javascript/api/word/word.section#getNextOrNullObject__)|Gets the next section.|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[getFirst()](/javascript/api/word/word.sectioncollection#getFirst__)|Gets the first section in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.sectioncollection#getFirstOrNullObject__)|Gets the first section in this collection.|
|[Table](/javascript/api/word/word.table)|[addColumns(insertLocation: Word.InsertLocation, columnCount: number, values?: string[][])](/javascript/api/word/word.table#addColumns_insertLocation__columnCount__values_)|Adds columns to the start or end of the table, using the first or last existing column as a template.|
||[addRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[][])](/javascript/api/word/word.table#addRows_insertLocation__rowCount__values_)|Adds rows to the start or end of the table, using the first or last existing row as a template.|
||[alignment](/javascript/api/word/word.table#alignment)|Gets or sets the alignment of the table against the page column.|
||[autoFitWindow()](/javascript/api/word/word.table#autoFitWindow__)|Autofits the table columns to the width of the window.|
||[clear()](/javascript/api/word/word.table#clear__)|Clears the contents of the table.|
||[delete()](/javascript/api/word/word.table#delete__)|Deletes the entire table.|
||[deleteColumns(columnIndex: number, columnCount?: number)](/javascript/api/word/word.table#deleteColumns_columnIndex__columnCount_)|Deletes specific columns.|
||[deleteRows(rowIndex: number, rowCount?: number)](/javascript/api/word/word.table#deleteRows_rowIndex__rowCount_)|Deletes specific rows.|
||[distributeColumns()](/javascript/api/word/word.table#distributeColumns__)|Distributes the column widths evenly.|
||[getBorder(borderLocation: Word.BorderLocation)](/javascript/api/word/word.table#getBorder_borderLocation_)|Gets the border style for the specified border.|
||[getCell(rowIndex: number, cellIndex: number)](/javascript/api/word/word.table#getCell_rowIndex__cellIndex_)|Gets the table cell at a specified row and column.|
||[getCellOrNullObject(rowIndex: number, cellIndex: number)](/javascript/api/word/word.table#getCellOrNullObject_rowIndex__cellIndex_)|Gets the table cell at a specified row and column.|
||[getCellPadding(cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.table#getCellPadding_cellPaddingLocation_)|Gets cell padding in points.|
||[getNext()](/javascript/api/word/word.table#getNext__)|Gets the next table.|
||[getNextOrNullObject()](/javascript/api/word/word.table#getNextOrNullObject__)|Gets the next table.|
||[getParagraphAfter()](/javascript/api/word/word.table#getParagraphAfter__)|Gets the paragraph after the table.|
||[getParagraphAfterOrNullObject()](/javascript/api/word/word.table#getParagraphAfterOrNullObject__)|Gets the paragraph after the table.|
||[getParagraphBefore()](/javascript/api/word/word.table#getParagraphBefore__)|Gets the paragraph before the table.|
||[getParagraphBeforeOrNullObject()](/javascript/api/word/word.table#getParagraphBeforeOrNullObject__)|Gets the paragraph before the table.|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.table#getRange_rangeLocation_)|Gets the range that contains this table, or the range at the start or end of the table.|
||[headerRowCount](/javascript/api/word/word.table#headerRowCount)|Gets and sets the number of header rows.|
||[horizontalAlignment](/javascript/api/word/word.table#horizontalAlignment)|Gets and sets the horizontal alignment of every cell in the table.|
||[ignorePunct](/javascript/api/word/word.table#ignorePunct)||
||[ignoreSpace](/javascript/api/word/word.table#ignoreSpace)||
||[insertContentControl()](/javascript/api/word/word.table#insertContentControl__)|Inserts a content control on the table.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.table#insertParagraph_paragraphText__insertLocation_)|Inserts a paragraph at the specified location.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.table#insertTable_rowCount__columnCount__insertLocation__values_)|Inserts a table with the specified number of rows and columns.|
||[matchCase](/javascript/api/word/word.table#matchCase)||
||[matchPrefix](/javascript/api/word/word.table#matchPrefix)||
||[matchSuffix](/javascript/api/word/word.table#matchSuffix)||
||[matchWholeWord](/javascript/api/word/word.table#matchWholeWord)||
||[matchWildcards](/javascript/api/word/word.table#matchWildcards)||
||[font](/javascript/api/word/word.table#font)|Gets the font.|
||[isUniform](/javascript/api/word/word.table#isUniform)|Indicates whether all of the table rows are uniform.|
||[nestingLevel](/javascript/api/word/word.table#nestingLevel)|Gets the nesting level of the table.|
||[parentBody](/javascript/api/word/word.table#parentBody)|Gets the parent body of the table.|
||[parentContentControl](/javascript/api/word/word.table#parentContentControl)|Gets the content control that contains the table.|
||[parentContentControlOrNullObject](/javascript/api/word/word.table#parentContentControlOrNullObject)|Gets the content control that contains the table.|
||[parentTable](/javascript/api/word/word.table#parentTable)|Gets the table that contains this table.|
||[parentTableCell](/javascript/api/word/word.table#parentTableCell)|Gets the table cell that contains this table.|
||[parentTableCellOrNullObject](/javascript/api/word/word.table#parentTableCellOrNullObject)|Gets the table cell that contains this table.|
||[parentTableOrNullObject](/javascript/api/word/word.table#parentTableOrNullObject)|Gets the table that contains this table.|
||[rowCount](/javascript/api/word/word.table#rowCount)|Gets the number of rows in the table.|
||[rows](/javascript/api/word/word.table#rows)|Gets all of the table rows.|
||[tables](/javascript/api/word/word.table#tables)|Gets the child tables nested one level deeper.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| {            ignorePunct?: boolean            ignoreSpace?: boolean            matchCase?: boolean            matchPrefix?: boolean            matchSuffix?: boolean            matchWholeWord?: boolean            matchWildcards?: boolean        })](/javascript/api/word/word.table#search_searchText__searchOptions__ignorePunct__ignoreSpace__matchCase__matchPrefix__matchSuffix__matchWholeWord__matchWildcards_)|Performs a search with the specified SearchOptions on the scope of the table object.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.table#select_selectionMode_)|Selects the table, or the position at the start or end of the table, and navigates the Word UI to it.|
||[setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.table#setCellPadding_cellPaddingLocation__cellPadding_)|Sets cell padding in points.|
||[shadingColor](/javascript/api/word/word.table#shadingColor)|Gets and sets the shading color.|
||[style](/javascript/api/word/word.table#style)|Gets or sets the style name for the table.|
||[styleBandedColumns](/javascript/api/word/word.table#styleBandedColumns)|Gets and sets whether the table has banded columns.|
||[styleBandedRows](/javascript/api/word/word.table#styleBandedRows)|Gets and sets whether the table has banded rows.|
||[styleBuiltIn](/javascript/api/word/word.table#styleBuiltIn)|Gets or sets the built-in style name for the table.|
||[styleFirstColumn](/javascript/api/word/word.table#styleFirstColumn)|Gets and sets whether the table has a first column with a special style.|
||[styleLastColumn](/javascript/api/word/word.table#styleLastColumn)|Gets and sets whether the table has a last column with a special style.|
||[styleTotalRow](/javascript/api/word/word.table#styleTotalRow)|Gets and sets whether the table has a total (last) row with a special style.|
||[values](/javascript/api/word/word.table#values)|Gets and sets the text values in the table, as a 2D Javascript array.|
||[verticalAlignment](/javascript/api/word/word.table#verticalAlignment)|Gets and sets the vertical alignment of every cell in the table.|
||[width](/javascript/api/word/word.table#width)|Gets and sets the width of the table in points.|
|[TableBorder](/javascript/api/word/word.tableborder)|[color](/javascript/api/word/word.tableborder#color)|Gets or sets the table border color.|
||[type](/javascript/api/word/word.tableborder#type)|Gets or sets the type of the table border.|
||[width](/javascript/api/word/word.tableborder#width)|Gets or sets the width, in points, of the table border.|
|[TableCell](/javascript/api/word/word.tablecell)|[columnWidth](/javascript/api/word/word.tablecell#columnWidth)|Gets and sets the width of the cell's column in points.|
||[deleteColumn()](/javascript/api/word/word.tablecell#deleteColumn__)|Deletes the column containing this cell.|
||[deleteRow()](/javascript/api/word/word.tablecell#deleteRow__)|Deletes the row containing this cell.|
||[getBorder(borderLocation: Word.BorderLocation)](/javascript/api/word/word.tablecell#getBorder_borderLocation_)|Gets the border style for the specified border.|
||[getCellPadding(cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.tablecell#getCellPadding_cellPaddingLocation_)|Gets cell padding in points.|
||[getNext()](/javascript/api/word/word.tablecell#getNext__)|Gets the next cell.|
||[getNextOrNullObject()](/javascript/api/word/word.tablecell#getNextOrNullObject__)|Gets the next cell.|
||[horizontalAlignment](/javascript/api/word/word.tablecell#horizontalAlignment)|Gets and sets the horizontal alignment of the cell.|
||[insertColumns(insertLocation: Word.InsertLocation, columnCount: number, values?: string[][])](/javascript/api/word/word.tablecell#insertColumns_insertLocation__columnCount__values_)|Adds columns to the left or right of the cell, using the cell's column as a template.|
||[insertRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[][])](/javascript/api/word/word.tablecell#insertRows_insertLocation__rowCount__values_)|Inserts rows above or below the cell, using the cell's row as a template.|
||[body](/javascript/api/word/word.tablecell#body)|Gets the body object of the cell.|
||[cellIndex](/javascript/api/word/word.tablecell#cellIndex)|Gets the index of the cell in its row.|
||[parentRow](/javascript/api/word/word.tablecell#parentRow)|Gets the parent row of the cell.|
||[parentTable](/javascript/api/word/word.tablecell#parentTable)|Gets the parent table of the cell.|
||[rowIndex](/javascript/api/word/word.tablecell#rowIndex)|Gets the index of the cell's row in the table.|
||[width](/javascript/api/word/word.tablecell#width)|Gets the width of the cell in points.|
||[setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.tablecell#setCellPadding_cellPaddingLocation__cellPadding_)|Sets cell padding in points.|
||[shadingColor](/javascript/api/word/word.tablecell#shadingColor)|Gets or sets the shading color of the cell.|
||[value](/javascript/api/word/word.tablecell#value)|Gets and sets the text of the cell.|
||[verticalAlignment](/javascript/api/word/word.tablecell#verticalAlignment)|Gets and sets the vertical alignment of the cell.|
|[TableCellCollection](/javascript/api/word/word.tablecellcollection)|[getFirst()](/javascript/api/word/word.tablecellcollection#getFirst__)|Gets the first table cell in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.tablecellcollection#getFirstOrNullObject__)|Gets the first table cell in this collection.|
||[items](/javascript/api/word/word.tablecellcollection#items)|Gets the loaded child items in this collection.|
|[TableCollection](/javascript/api/word/word.tablecollection)|[getFirst()](/javascript/api/word/word.tablecollection#getFirst__)|Gets the first table in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.tablecollection#getFirstOrNullObject__)|Gets the first table in this collection.|
||[items](/javascript/api/word/word.tablecollection#items)|Gets the loaded child items in this collection.|
|[TableRow](/javascript/api/word/word.tablerow)|[clear()](/javascript/api/word/word.tablerow#clear__)|Clears the contents of the row.|
||[delete()](/javascript/api/word/word.tablerow#delete__)|Deletes the entire row.|
||[getBorder(borderLocation: Word.BorderLocation)](/javascript/api/word/word.tablerow#getBorder_borderLocation_)|Gets the border style of the cells in the row.|
||[getCellPadding(cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.tablerow#getCellPadding_cellPaddingLocation_)|Gets cell padding in points.|
||[getNext()](/javascript/api/word/word.tablerow#getNext__)|Gets the next row.|
||[getNextOrNullObject()](/javascript/api/word/word.tablerow#getNextOrNullObject__)|Gets the next row.|
||[horizontalAlignment](/javascript/api/word/word.tablerow#horizontalAlignment)|Gets and sets the horizontal alignment of every cell in the row.|
||[ignorePunct](/javascript/api/word/word.tablerow#ignorePunct)||
||[ignoreSpace](/javascript/api/word/word.tablerow#ignoreSpace)||
||[insertRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[][])](/javascript/api/word/word.tablerow#insertRows_insertLocation__rowCount__values_)|Inserts rows using this row as a template.|
||[matchCase](/javascript/api/word/word.tablerow#matchCase)||
||[matchPrefix](/javascript/api/word/word.tablerow#matchPrefix)||
||[matchSuffix](/javascript/api/word/word.tablerow#matchSuffix)||
||[matchWholeWord](/javascript/api/word/word.tablerow#matchWholeWord)||
||[matchWildcards](/javascript/api/word/word.tablerow#matchWildcards)||
||[preferredHeight](/javascript/api/word/word.tablerow#preferredHeight)|Gets and sets the preferred height of the row in points.|
||[cellCount](/javascript/api/word/word.tablerow#cellCount)|Gets the number of cells in the row.|
||[cells](/javascript/api/word/word.tablerow#cells)|Gets cells.|
||[font](/javascript/api/word/word.tablerow#font)|Gets the font.|
||[isHeader](/javascript/api/word/word.tablerow#isHeader)|Checks whether the row is a header row.|
||[parentTable](/javascript/api/word/word.tablerow#parentTable)|Gets parent table.|
||[rowIndex](/javascript/api/word/word.tablerow#rowIndex)|Gets the index of the row in its parent table.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| {            ignorePunct?: boolean            ignoreSpace?: boolean            matchCase?: boolean            matchPrefix?: boolean            matchSuffix?: boolean            matchWholeWord?: boolean            matchWildcards?: boolean        })](/javascript/api/word/word.tablerow#search_searchText__searchOptions__ignorePunct__ignoreSpace__matchCase__matchPrefix__matchSuffix__matchWholeWord__matchWildcards_)|Performs a search with the specified SearchOptions on the scope of the row.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.tablerow#select_selectionMode_)|Selects the row and navigates the Word UI to it.|
||[setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.tablerow#setCellPadding_cellPaddingLocation__cellPadding_)|Sets cell padding in points.|
||[shadingColor](/javascript/api/word/word.tablerow#shadingColor)|Gets and sets the shading color.|
||[values](/javascript/api/word/word.tablerow#values)|Gets and sets the text values in the row, as a 2D Javascript array.|
||[verticalAlignment](/javascript/api/word/word.tablerow#verticalAlignment)|Gets and sets the vertical alignment of the cells in the row.|
|[TableRowCollection](/javascript/api/word/word.tablerowcollection)|[getFirst()](/javascript/api/word/word.tablerowcollection#getFirst__)|Gets the first row in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.tablerowcollection#getFirstOrNullObject__)|Gets the first row in this collection.|
||[items](/javascript/api/word/word.tablerowcollection#items)|Gets the loaded child items in this collection.|

## See also

- [Word JavaScript API Reference Documentation](/javascript/api/word)
- [Word JavaScript API requirement sets](word-api-requirement-sets.md)
