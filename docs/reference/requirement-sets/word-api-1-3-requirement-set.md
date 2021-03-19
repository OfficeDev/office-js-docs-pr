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
|[Application](/javascript/api/word/word.application)|[createDocument(base64File?: string)](/javascript/api/word/word.application#createdocument-base64file-)|Creates a new document by using an optional base64 encoded .docx file.|
|[Body](/javascript/api/word/word.body)|[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.body#getrange-rangelocation-)|Gets the whole body, or the starting or ending point of the body, as a range.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.body#inserttable-rowcount--columncount--insertlocation--values-)|Inserts a table with the specified number of rows and columns.|
||[lists](/javascript/api/word/word.body#lists)|Gets the collection of list objects in the body.|
||[parentBody](/javascript/api/word/word.body#parentbody)|Gets the parent body of the body.|
||[parentBodyOrNullObject](/javascript/api/word/word.body#parentbodyornullobject)|Gets the parent body of the body.|
||[parentContentControlOrNullObject](/javascript/api/word/word.body#parentcontentcontrolornullobject)|Gets the content control that contains the body.|
||[parentSection](/javascript/api/word/word.body#parentsection)|Gets the parent section of the body.|
||[parentSectionOrNullObject](/javascript/api/word/word.body#parentsectionornullobject)|Gets the parent section of the body.|
||[tables](/javascript/api/word/word.body#tables)|Gets the collection of table objects in the body.|
||[type](/javascript/api/word/word.body#type)|Gets the type of the body.|
||[styleBuiltIn](/javascript/api/word/word.body#stylebuiltin)|Gets or sets the built-in style name for the body.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.contentcontrol#getrange-rangelocation-)|Gets the whole content control, or the starting or ending point of the content control, as a range.|
||[getTextRanges(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.contentcontrol#gettextranges-endingmarks--trimspacing-)|Gets the text ranges in the content control by using punctuation marks and/or other ending marks.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.contentcontrol#inserttable-rowcount--columncount--insertlocation--values-)|Inserts a table with the specified number of rows and columns into, or next to, a content control.|
||[lists](/javascript/api/word/word.contentcontrol#lists)|Gets the collection of list objects in the content control.|
||[parentBody](/javascript/api/word/word.contentcontrol#parentbody)|Gets the parent body of the content control.|
||[parentContentControlOrNullObject](/javascript/api/word/word.contentcontrol#parentcontentcontrolornullobject)|Gets the content control that contains the content control.|
||[parentTable](/javascript/api/word/word.contentcontrol#parenttable)|Gets the table that contains the content control.|
||[parentTableCell](/javascript/api/word/word.contentcontrol#parenttablecell)|Gets the table cell that contains the content control.|
||[parentTableCellOrNullObject](/javascript/api/word/word.contentcontrol#parenttablecellornullobject)|Gets the table cell that contains the content control.|
||[parentTableOrNullObject](/javascript/api/word/word.contentcontrol#parenttableornullobject)|Gets the table that contains the content control.|
||[subtype](/javascript/api/word/word.contentcontrol#subtype)|Gets the content control subtype.|
||[tables](/javascript/api/word/word.contentcontrol#tables)|Gets the collection of table objects in the content control.|
||[split(delimiters: string[], multiParagraphs?: boolean, trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.contentcontrol#split-delimiters--multiparagraphs--trimdelimiters--trimspacing-)|Splits the content control into child ranges by using delimiters.|
||[styleBuiltIn](/javascript/api/word/word.contentcontrol#stylebuiltin)|Gets or sets the built-in style name for the content control.|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getByIdOrNullObject(id: number)](/javascript/api/word/word.contentcontrolcollection#getbyidornullobject-id-)|Gets a content control by its identifier.|
||[getByTypes(types: Word.ContentControlType[])](/javascript/api/word/word.contentcontrolcollection#getbytypes-types-)|Gets the content controls that have the specified types and/or subtypes.|
||[getFirst()](/javascript/api/word/word.contentcontrolcollection#getfirst--)|Gets the first content control in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.contentcontrolcollection#getfirstornullobject--)|Gets the first content control in this collection.|
|[CustomProperty](/javascript/api/word/word.customproperty)|[delete()](/javascript/api/word/word.customproperty#delete--)|Deletes the custom property.|
||[key](/javascript/api/word/word.customproperty#key)|Gets the key of the custom property.|
||[type](/javascript/api/word/word.customproperty#type)|Gets the value type of the custom property.|
||[value](/javascript/api/word/word.customproperty#value)|Gets or sets the value of the custom property.|
|[CustomPropertyCollection](/javascript/api/word/word.custompropertycollection)|[add(key: string, value: any)](/javascript/api/word/word.custompropertycollection#add-key--value-)|Creates a new or sets an existing custom property.|
||[deleteAll()](/javascript/api/word/word.custompropertycollection#deleteall--)|Deletes all custom properties in this collection.|
||[getCount()](/javascript/api/word/word.custompropertycollection#getcount--)|Gets the count of custom properties.|
||[getItem(key: string)](/javascript/api/word/word.custompropertycollection#getitem-key-)|Gets a custom property object by its key, which is case-insensitive.|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.custompropertycollection#getitemornullobject-key-)|Gets a custom property object by its key, which is case-insensitive.|
||[items](/javascript/api/word/word.custompropertycollection#items)|Gets the loaded child items in this collection.|
|[Document](/javascript/api/word/word.document)|[properties](/javascript/api/word/word.document#properties)|Gets the properties of the document.|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[open()](/javascript/api/word/word.documentcreated#open--)|Opens the document.|
||[body](/javascript/api/word/word.documentcreated#body)|Gets the body object of the document.|
||[contentControls](/javascript/api/word/word.documentcreated#contentcontrols)|Gets the collection of content control objects in the document.|
||[properties](/javascript/api/word/word.documentcreated#properties)|Gets the properties of the document.|
||[saved](/javascript/api/word/word.documentcreated#saved)|Indicates whether the changes in the document have been saved.|
||[sections](/javascript/api/word/word.documentcreated#sections)|Gets the collection of section objects in the document.|
||[save()](/javascript/api/word/word.documentcreated#save--)|Saves the document.|
|[DocumentProperties](/javascript/api/word/word.documentproperties)|[author](/javascript/api/word/word.documentproperties#author)|Gets or sets the author of the document.|
||[category](/javascript/api/word/word.documentproperties#category)|Gets or sets the category of the document.|
||[comments](/javascript/api/word/word.documentproperties#comments)|Gets or sets the comments of the document.|
||[company](/javascript/api/word/word.documentproperties#company)|Gets or sets the company of the document.|
||[format](/javascript/api/word/word.documentproperties#format)|Gets or sets the format of the document.|
||[keywords](/javascript/api/word/word.documentproperties#keywords)|Gets or sets the keywords of the document.|
||[manager](/javascript/api/word/word.documentproperties#manager)|Gets or sets the manager of the document.|
||[applicationName](/javascript/api/word/word.documentproperties#applicationname)|Gets the application name of the document.|
||[creationDate](/javascript/api/word/word.documentproperties#creationdate)|Gets the creation date of the document.|
||[customProperties](/javascript/api/word/word.documentproperties#customproperties)|Gets the collection of custom properties of the document.|
||[lastAuthor](/javascript/api/word/word.documentproperties#lastauthor)|Gets the last author of the document.|
||[lastPrintDate](/javascript/api/word/word.documentproperties#lastprintdate)|Gets the last print date of the document.|
||[lastSaveTime](/javascript/api/word/word.documentproperties#lastsavetime)|Gets the last save time of the document.|
||[revisionNumber](/javascript/api/word/word.documentproperties#revisionnumber)|Gets the revision number of the document.|
||[security](/javascript/api/word/word.documentproperties#security)|Gets security settings of the document.|
||[template](/javascript/api/word/word.documentproperties#template)|Gets the template of the document.|
||[subject](/javascript/api/word/word.documentproperties#subject)|Gets or sets the subject of the document.|
||[title](/javascript/api/word/word.documentproperties#title)|Gets or sets the title of the document.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[getNext()](/javascript/api/word/word.inlinepicture#getnext--)|Gets the next inline image.|
||[getNextOrNullObject()](/javascript/api/word/word.inlinepicture#getnextornullobject--)|Gets the next inline image.|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.inlinepicture#getrange-rangelocation-)|Gets the picture, or the starting or ending point of the picture, as a range.|
||[parentContentControlOrNullObject](/javascript/api/word/word.inlinepicture#parentcontentcontrolornullobject)|Gets the content control that contains the inline image.|
||[parentTable](/javascript/api/word/word.inlinepicture#parenttable)|Gets the table that contains the inline image.|
||[parentTableCell](/javascript/api/word/word.inlinepicture#parenttablecell)|Gets the table cell that contains the inline image.|
||[parentTableCellOrNullObject](/javascript/api/word/word.inlinepicture#parenttablecellornullobject)|Gets the table cell that contains the inline image.|
||[parentTableOrNullObject](/javascript/api/word/word.inlinepicture#parenttableornullobject)|Gets the table that contains the inline image.|
|[InlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|[getFirst()](/javascript/api/word/word.inlinepicturecollection#getfirst--)|Gets the first inline image in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.inlinepicturecollection#getfirstornullobject--)|Gets the first inline image in this collection.|
|[List](/javascript/api/word/word.list)|[getLevelParagraphs(level: number)](/javascript/api/word/word.list#getlevelparagraphs-level-)|Gets the paragraphs that occur at the specified level in the list.|
||[getLevelString(level: number)](/javascript/api/word/word.list#getlevelstring-level-)|Gets the bullet, number, or picture at the specified level as a string.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.list#insertparagraph-paragraphtext--insertlocation-)|Inserts a paragraph at the specified location.|
||[id](/javascript/api/word/word.list#id)|Gets the list's id.|
||[levelExistences](/javascript/api/word/word.list#levelexistences)|Checks whether each of the 9 levels exists in the list.|
||[levelTypes](/javascript/api/word/word.list#leveltypes)|Gets all 9 level types in the list.|
||[paragraphs](/javascript/api/word/word.list#paragraphs)|Gets paragraphs in the list.|
||[setLevelAlignment(level: number, alignment: Word.Alignment)](/javascript/api/word/word.list#setlevelalignment-level--alignment-)|Sets the alignment of the bullet, number, or picture at the specified level in the list.|
||[setLevelBullet(level: number, listBullet: Word.ListBullet, charCode?: number, fontName?: string)](/javascript/api/word/word.list#setlevelbullet-level--listbullet--charcode--fontname-)|Sets the bullet format at the specified level in the list.|
||[setLevelIndents(level: number, textIndent: number, bulletNumberPictureIndent: number)](/javascript/api/word/word.list#setlevelindents-level--textindent--bulletnumberpictureindent-)|Sets the two indents of the specified level in the list.|
||[setLevelNumbering(level: number, listNumbering: Word.ListNumbering, formatString?: Array<string \| number>)](/javascript/api/word/word.list#setlevelnumbering-level--listnumbering--formatstring-)|Sets the numbering format at the specified level in the list.|
||[setLevelStartingNumber(level: number, startingNumber: number)](/javascript/api/word/word.list#setlevelstartingnumber-level--startingnumber-)|Sets the starting number at the specified level in the list.|
|[ListCollection](/javascript/api/word/word.listcollection)|[getById(id: number)](/javascript/api/word/word.listcollection#getbyid-id-)|Gets a list by its identifier.|
||[getByIdOrNullObject(id: number)](/javascript/api/word/word.listcollection#getbyidornullobject-id-)|Gets a list by its identifier.|
||[getFirst()](/javascript/api/word/word.listcollection#getfirst--)|Gets the first list in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.listcollection#getfirstornullobject--)|Gets the first list in this collection.|
||[getItem(index: number)](/javascript/api/word/word.listcollection#getitem-index-)|Gets a list object by its index in the collection.|
||[items](/javascript/api/word/word.listcollection#items)|Gets the loaded child items in this collection.|
|[ListItem](/javascript/api/word/word.listitem)|[getAncestor(parentOnly?: boolean)](/javascript/api/word/word.listitem#getancestor-parentonly-)|Gets the list item parent, or the closest ancestor if the parent does not exist.|
||[getAncestorOrNullObject(parentOnly?: boolean)](/javascript/api/word/word.listitem#getancestorornullobject-parentonly-)|Gets the list item parent, or the closest ancestor if the parent does not exist.|
||[getDescendants(directChildrenOnly?: boolean)](/javascript/api/word/word.listitem#getdescendants-directchildrenonly-)|Gets all descendant list items of the list item.|
||[level](/javascript/api/word/word.listitem#level)|Gets or sets the level of the item in the list.|
||[listString](/javascript/api/word/word.listitem#liststring)|Gets the list item bullet, number, or picture as a string.|
||[siblingIndex](/javascript/api/word/word.listitem#siblingindex)|Gets the list item order number in relation to its siblings.|
|[Paragraph](/javascript/api/word/word.paragraph)|[attachToList(listId: number, level: number)](/javascript/api/word/word.paragraph#attachtolist-listid--level-)|Lets the paragraph join an existing list at the specified level.|
||[detachFromList()](/javascript/api/word/word.paragraph#detachfromlist--)|Moves this paragraph out of its list, if the paragraph is a list item.|
||[getNext()](/javascript/api/word/word.paragraph#getnext--)|Gets the next paragraph.|
||[getNextOrNullObject()](/javascript/api/word/word.paragraph#getnextornullobject--)|Gets the next paragraph.|
||[getPrevious()](/javascript/api/word/word.paragraph#getprevious--)|Gets the previous paragraph.|
||[getPreviousOrNullObject()](/javascript/api/word/word.paragraph#getpreviousornullobject--)|Gets the previous paragraph.|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.paragraph#getrange-rangelocation-)|Gets the whole paragraph, or the starting or ending point of the paragraph, as a range.|
||[getTextRanges(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.paragraph#gettextranges-endingmarks--trimspacing-)|Gets the text ranges in the paragraph by using punctuation marks and/or other ending marks.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.paragraph#inserttable-rowcount--columncount--insertlocation--values-)|Inserts a table with the specified number of rows and columns.|
||[isLastParagraph](/javascript/api/word/word.paragraph#islastparagraph)|Indicates the paragraph is the last one inside its parent body.|
||[isListItem](/javascript/api/word/word.paragraph#islistitem)|Checks whether the paragraph is a list item.|
||[list](/javascript/api/word/word.paragraph#list)|Gets the List to which this paragraph belongs.|
||[listItem](/javascript/api/word/word.paragraph#listitem)|Gets the ListItem for the paragraph.|
||[listItemOrNullObject](/javascript/api/word/word.paragraph#listitemornullobject)|Gets the ListItem for the paragraph.|
||[listOrNullObject](/javascript/api/word/word.paragraph#listornullobject)|Gets the List to which this paragraph belongs.|
||[parentBody](/javascript/api/word/word.paragraph#parentbody)|Gets the parent body of the paragraph.|
||[parentContentControlOrNullObject](/javascript/api/word/word.paragraph#parentcontentcontrolornullobject)|Gets the content control that contains the paragraph.|
||[parentTable](/javascript/api/word/word.paragraph#parenttable)|Gets the table that contains the paragraph.|
||[parentTableCell](/javascript/api/word/word.paragraph#parenttablecell)|Gets the table cell that contains the paragraph.|
||[parentTableCellOrNullObject](/javascript/api/word/word.paragraph#parenttablecellornullobject)|Gets the table cell that contains the paragraph.|
||[parentTableOrNullObject](/javascript/api/word/word.paragraph#parenttableornullobject)|Gets the table that contains the paragraph.|
||[tableNestingLevel](/javascript/api/word/word.paragraph#tablenestinglevel)|Gets the level of the paragraph's table.|
||[split(delimiters: string[], trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.paragraph#split-delimiters--trimdelimiters--trimspacing-)|Splits the paragraph into child ranges by using delimiters.|
||[startNewList()](/javascript/api/word/word.paragraph#startnewlist--)|Starts a new list with this paragraph.|
||[styleBuiltIn](/javascript/api/word/word.paragraph#stylebuiltin)|Gets or sets the built-in style name for the paragraph.|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[getFirst()](/javascript/api/word/word.paragraphcollection#getfirst--)|Gets the first paragraph in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.paragraphcollection#getfirstornullobject--)|Gets the first paragraph in this collection.|
||[getLast()](/javascript/api/word/word.paragraphcollection#getlast--)|Gets the last paragraph in this collection.|
||[getLastOrNullObject()](/javascript/api/word/word.paragraphcollection#getlastornullobject--)|Gets the last paragraph in this collection.|
|[Range](/javascript/api/word/word.range)|[compareLocationWith(range: Word.Range)](/javascript/api/word/word.range#comparelocationwith-range-)|Compares this range's location with another range's location.|
||[expandTo(range: Word.Range)](/javascript/api/word/word.range#expandto-range-)|Returns a new range that extends from this range in either direction to cover another range.|
||[expandToOrNullObject(range: Word.Range)](/javascript/api/word/word.range#expandtoornullobject-range-)|Returns a new range that extends from this range in either direction to cover another range.|
||[getHyperlinkRanges()](/javascript/api/word/word.range#gethyperlinkranges--)|Gets hyperlink child ranges within the range.|
||[getNextTextRange(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#getnexttextrange-endingmarks--trimspacing-)|Gets the next text range by using punctuation marks and/or other ending marks.|
||[getNextTextRangeOrNullObject(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#getnexttextrangeornullobject-endingmarks--trimspacing-)|Gets the next text range by using punctuation marks and/or other ending marks.|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.range#getrange-rangelocation-)|Clones the range, or gets the starting or ending point of the range as a new range.|
||[getTextRanges(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#gettextranges-endingmarks--trimspacing-)|Gets the text child ranges in the range by using punctuation marks and/or other ending marks.|
||[hyperlink](/javascript/api/word/word.range#hyperlink)|Gets the first hyperlink in the range, or sets a hyperlink on the range.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.range#inserttable-rowcount--columncount--insertlocation--values-)|Inserts a table with the specified number of rows and columns.|
||[intersectWith(range: Word.Range)](/javascript/api/word/word.range#intersectwith-range-)|Returns a new range as the intersection of this range with another range.|
||[intersectWithOrNullObject(range: Word.Range)](/javascript/api/word/word.range#intersectwithornullobject-range-)|Returns a new range as the intersection of this range with another range.|
||[isEmpty](/javascript/api/word/word.range#isempty)|Checks whether the range length is zero.|
||[lists](/javascript/api/word/word.range#lists)|Gets the collection of list objects in the range.|
||[parentBody](/javascript/api/word/word.range#parentbody)|Gets the parent body of the range.|
||[parentContentControlOrNullObject](/javascript/api/word/word.range#parentcontentcontrolornullobject)|Gets the content control that contains the range.|
||[parentTable](/javascript/api/word/word.range#parenttable)|Gets the table that contains the range.|
||[parentTableCell](/javascript/api/word/word.range#parenttablecell)|Gets the table cell that contains the range.|
||[parentTableCellOrNullObject](/javascript/api/word/word.range#parenttablecellornullobject)|Gets the table cell that contains the range.|
||[parentTableOrNullObject](/javascript/api/word/word.range#parenttableornullobject)|Gets the table that contains the range.|
||[tables](/javascript/api/word/word.range#tables)|Gets the collection of table objects in the range.|
||[split(delimiters: string[], multiParagraphs?: boolean, trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.range#split-delimiters--multiparagraphs--trimdelimiters--trimspacing-)|Splits the range into child ranges by using delimiters.|
||[styleBuiltIn](/javascript/api/word/word.range#stylebuiltin)|Gets or sets the built-in style name for the range.|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[getFirst()](/javascript/api/word/word.rangecollection#getfirst--)|Gets the first range in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.rangecollection#getfirstornullobject--)|Gets the first range in this collection.|
|[RequestContext](/javascript/api/word/word.requestcontext)|[application](/javascript/api/word/word.requestcontext#application)|[Api set: WordApi 1.3] *|
|[Section](/javascript/api/word/word.section)|[getNext()](/javascript/api/word/word.section#getnext--)|Gets the next section.|
||[getNextOrNullObject()](/javascript/api/word/word.section#getnextornullobject--)|Gets the next section.|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[getFirst()](/javascript/api/word/word.sectioncollection#getfirst--)|Gets the first section in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.sectioncollection#getfirstornullobject--)|Gets the first section in this collection.|
|[Table](/javascript/api/word/word.table)|[addColumns(insertLocation: Word.InsertLocation, columnCount: number, values?: string[][])](/javascript/api/word/word.table#addcolumns-insertlocation--columncount--values-)|Adds columns to the start or end of the table, using the first or last existing column as a template.|
||[addRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[][])](/javascript/api/word/word.table#addrows-insertlocation--rowcount--values-)|Adds rows to the start or end of the table, using the first or last existing row as a template.|
||[alignment](/javascript/api/word/word.table#alignment)|Gets or sets the alignment of the table against the page column.|
||[autoFitWindow()](/javascript/api/word/word.table#autofitwindow--)|Autofits the table columns to the width of the window.|
||[clear()](/javascript/api/word/word.table#clear--)|Clears the contents of the table.|
||[delete()](/javascript/api/word/word.table#delete--)|Deletes the entire table.|
||[deleteColumns(columnIndex: number, columnCount?: number)](/javascript/api/word/word.table#deletecolumns-columnindex--columncount-)|Deletes specific columns.|
||[deleteRows(rowIndex: number, rowCount?: number)](/javascript/api/word/word.table#deleterows-rowindex--rowcount-)|Deletes specific rows.|
||[distributeColumns()](/javascript/api/word/word.table#distributecolumns--)|Distributes the column widths evenly.|
||[getBorder(borderLocation: Word.BorderLocation)](/javascript/api/word/word.table#getborder-borderlocation-)|Gets the border style for the specified border.|
||[getCell(rowIndex: number, cellIndex: number)](/javascript/api/word/word.table#getcell-rowindex--cellindex-)|Gets the table cell at a specified row and column.|
||[getCellOrNullObject(rowIndex: number, cellIndex: number)](/javascript/api/word/word.table#getcellornullobject-rowindex--cellindex-)|Gets the table cell at a specified row and column.|
||[getCellPadding(cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.table#getcellpadding-cellpaddinglocation-)|Gets cell padding in points.|
||[getNext()](/javascript/api/word/word.table#getnext--)|Gets the next table.|
||[getNextOrNullObject()](/javascript/api/word/word.table#getnextornullobject--)|Gets the next table.|
||[getParagraphAfter()](/javascript/api/word/word.table#getparagraphafter--)|Gets the paragraph after the table.|
||[getParagraphAfterOrNullObject()](/javascript/api/word/word.table#getparagraphafterornullobject--)|Gets the paragraph after the table.|
||[getParagraphBefore()](/javascript/api/word/word.table#getparagraphbefore--)|Gets the paragraph before the table.|
||[getParagraphBeforeOrNullObject()](/javascript/api/word/word.table#getparagraphbeforeornullobject--)|Gets the paragraph before the table.|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.table#getrange-rangelocation-)|Gets the range that contains this table, or the range at the start or end of the table.|
||[headerRowCount](/javascript/api/word/word.table#headerrowcount)|Gets and sets the number of header rows.|
||[horizontalAlignment](/javascript/api/word/word.table#horizontalalignment)|Gets and sets the horizontal alignment of every cell in the table.|
||[ignorePunct](/javascript/api/word/word.table#ignorepunct)||
||[ignoreSpace](/javascript/api/word/word.table#ignorespace)||
||[insertContentControl()](/javascript/api/word/word.table#insertcontentcontrol--)|Inserts a content control on the table.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.table#insertparagraph-paragraphtext--insertlocation-)|Inserts a paragraph at the specified location.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.table#inserttable-rowcount--columncount--insertlocation--values-)|Inserts a table with the specified number of rows and columns.|
||[matchCase](/javascript/api/word/word.table#matchcase)||
||[matchPrefix](/javascript/api/word/word.table#matchprefix)||
||[matchSuffix](/javascript/api/word/word.table#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.table#matchwholeword)||
||[matchWildcards](/javascript/api/word/word.table#matchwildcards)||
||[font](/javascript/api/word/word.table#font)|Gets the font.|
||[isUniform](/javascript/api/word/word.table#isuniform)|Indicates whether all of the table rows are uniform.|
||[nestingLevel](/javascript/api/word/word.table#nestinglevel)|Gets the nesting level of the table.|
||[parentBody](/javascript/api/word/word.table#parentbody)|Gets the parent body of the table.|
||[parentContentControl](/javascript/api/word/word.table#parentcontentcontrol)|Gets the content control that contains the table.|
||[parentContentControlOrNullObject](/javascript/api/word/word.table#parentcontentcontrolornullobject)|Gets the content control that contains the table.|
||[parentTable](/javascript/api/word/word.table#parenttable)|Gets the table that contains this table.|
||[parentTableCell](/javascript/api/word/word.table#parenttablecell)|Gets the table cell that contains this table.|
||[parentTableCellOrNullObject](/javascript/api/word/word.table#parenttablecellornullobject)|Gets the table cell that contains this table.|
||[parentTableOrNullObject](/javascript/api/word/word.table#parenttableornullobject)|Gets the table that contains this table.|
||[rowCount](/javascript/api/word/word.table#rowcount)|Gets the number of rows in the table.|
||[rows](/javascript/api/word/word.table#rows)|Gets all of the table rows.|
||[tables](/javascript/api/word/word.table#tables)|Gets the child tables nested one level deeper.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| {            ignorePunct?: boolean            ignoreSpace?: boolean            matchCase?: boolean            matchPrefix?: boolean            matchSuffix?: boolean            matchWholeWord?: boolean            matchWildcards?: boolean        })](/javascript/api/word/word.table#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|Performs a search with the specified SearchOptions on the scope of the table object.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.table#select-selectionmode-)|Selects the table, or the position at the start or end of the table, and navigates the Word UI to it.|
||[setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.table#setcellpadding-cellpaddinglocation--cellpadding-)|Sets cell padding in points.|
||[shadingColor](/javascript/api/word/word.table#shadingcolor)|Gets and sets the shading color.|
||[style](/javascript/api/word/word.table#style)|Gets or sets the style name for the table.|
||[styleBandedColumns](/javascript/api/word/word.table#stylebandedcolumns)|Gets and sets whether the table has banded columns.|
||[styleBandedRows](/javascript/api/word/word.table#stylebandedrows)|Gets and sets whether the table has banded rows.|
||[styleBuiltIn](/javascript/api/word/word.table#stylebuiltin)|Gets or sets the built-in style name for the table.|
||[styleFirstColumn](/javascript/api/word/word.table#stylefirstcolumn)|Gets and sets whether the table has a first column with a special style.|
||[styleLastColumn](/javascript/api/word/word.table#stylelastcolumn)|Gets and sets whether the table has a last column with a special style.|
||[styleTotalRow](/javascript/api/word/word.table#styletotalrow)|Gets and sets whether the table has a total (last) row with a special style.|
||[values](/javascript/api/word/word.table#values)|Gets and sets the text values in the table, as a 2D Javascript array.|
||[verticalAlignment](/javascript/api/word/word.table#verticalalignment)|Gets and sets the vertical alignment of every cell in the table.|
||[width](/javascript/api/word/word.table#width)|Gets and sets the width of the table in points.|
|[TableBorder](/javascript/api/word/word.tableborder)|[color](/javascript/api/word/word.tableborder#color)|Gets or sets the table border color.|
||[type](/javascript/api/word/word.tableborder#type)|Gets or sets the type of the table border.|
||[width](/javascript/api/word/word.tableborder#width)|Gets or sets the width, in points, of the table border.|
|[TableCell](/javascript/api/word/word.tablecell)|[columnWidth](/javascript/api/word/word.tablecell#columnwidth)|Gets and sets the width of the cell's column in points.|
||[deleteColumn()](/javascript/api/word/word.tablecell#deletecolumn--)|Deletes the column containing this cell.|
||[deleteRow()](/javascript/api/word/word.tablecell#deleterow--)|Deletes the row containing this cell.|
||[getBorder(borderLocation: Word.BorderLocation)](/javascript/api/word/word.tablecell#getborder-borderlocation-)|Gets the border style for the specified border.|
||[getCellPadding(cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.tablecell#getcellpadding-cellpaddinglocation-)|Gets cell padding in points.|
||[getNext()](/javascript/api/word/word.tablecell#getnext--)|Gets the next cell.|
||[getNextOrNullObject()](/javascript/api/word/word.tablecell#getnextornullobject--)|Gets the next cell.|
||[horizontalAlignment](/javascript/api/word/word.tablecell#horizontalalignment)|Gets and sets the horizontal alignment of the cell.|
||[insertColumns(insertLocation: Word.InsertLocation, columnCount: number, values?: string[][])](/javascript/api/word/word.tablecell#insertcolumns-insertlocation--columncount--values-)|Adds columns to the left or right of the cell, using the cell's column as a template.|
||[insertRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[][])](/javascript/api/word/word.tablecell#insertrows-insertlocation--rowcount--values-)|Inserts rows above or below the cell, using the cell's row as a template.|
||[body](/javascript/api/word/word.tablecell#body)|Gets the body object of the cell.|
||[cellIndex](/javascript/api/word/word.tablecell#cellindex)|Gets the index of the cell in its row.|
||[parentRow](/javascript/api/word/word.tablecell#parentrow)|Gets the parent row of the cell.|
||[parentTable](/javascript/api/word/word.tablecell#parenttable)|Gets the parent table of the cell.|
||[rowIndex](/javascript/api/word/word.tablecell#rowindex)|Gets the index of the cell's row in the table.|
||[width](/javascript/api/word/word.tablecell#width)|Gets the width of the cell in points.|
||[setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.tablecell#setcellpadding-cellpaddinglocation--cellpadding-)|Sets cell padding in points.|
||[shadingColor](/javascript/api/word/word.tablecell#shadingcolor)|Gets or sets the shading color of the cell.|
||[value](/javascript/api/word/word.tablecell#value)|Gets and sets the text of the cell.|
||[verticalAlignment](/javascript/api/word/word.tablecell#verticalalignment)|Gets and sets the vertical alignment of the cell.|
|[TableCellCollection](/javascript/api/word/word.tablecellcollection)|[getFirst()](/javascript/api/word/word.tablecellcollection#getfirst--)|Gets the first table cell in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.tablecellcollection#getfirstornullobject--)|Gets the first table cell in this collection.|
||[items](/javascript/api/word/word.tablecellcollection#items)|Gets the loaded child items in this collection.|
|[TableCollection](/javascript/api/word/word.tablecollection)|[getFirst()](/javascript/api/word/word.tablecollection#getfirst--)|Gets the first table in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.tablecollection#getfirstornullobject--)|Gets the first table in this collection.|
||[items](/javascript/api/word/word.tablecollection#items)|Gets the loaded child items in this collection.|
|[TableRow](/javascript/api/word/word.tablerow)|[clear()](/javascript/api/word/word.tablerow#clear--)|Clears the contents of the row.|
||[delete()](/javascript/api/word/word.tablerow#delete--)|Deletes the entire row.|
||[getBorder(borderLocation: Word.BorderLocation)](/javascript/api/word/word.tablerow#getborder-borderlocation-)|Gets the border style of the cells in the row.|
||[getCellPadding(cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.tablerow#getcellpadding-cellpaddinglocation-)|Gets cell padding in points.|
||[getNext()](/javascript/api/word/word.tablerow#getnext--)|Gets the next row.|
||[getNextOrNullObject()](/javascript/api/word/word.tablerow#getnextornullobject--)|Gets the next row.|
||[horizontalAlignment](/javascript/api/word/word.tablerow#horizontalalignment)|Gets and sets the horizontal alignment of every cell in the row.|
||[ignorePunct](/javascript/api/word/word.tablerow#ignorepunct)||
||[ignoreSpace](/javascript/api/word/word.tablerow#ignorespace)||
||[insertRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[][])](/javascript/api/word/word.tablerow#insertrows-insertlocation--rowcount--values-)|Inserts rows using this row as a template.|
||[matchCase](/javascript/api/word/word.tablerow#matchcase)||
||[matchPrefix](/javascript/api/word/word.tablerow#matchprefix)||
||[matchSuffix](/javascript/api/word/word.tablerow#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.tablerow#matchwholeword)||
||[matchWildcards](/javascript/api/word/word.tablerow#matchwildcards)||
||[preferredHeight](/javascript/api/word/word.tablerow#preferredheight)|Gets and sets the preferred height of the row in points.|
||[cellCount](/javascript/api/word/word.tablerow#cellcount)|Gets the number of cells in the row.|
||[cells](/javascript/api/word/word.tablerow#cells)|Gets cells.|
||[font](/javascript/api/word/word.tablerow#font)|Gets the font.|
||[isHeader](/javascript/api/word/word.tablerow#isheader)|Checks whether the row is a header row.|
||[parentTable](/javascript/api/word/word.tablerow#parenttable)|Gets parent table.|
||[rowIndex](/javascript/api/word/word.tablerow#rowindex)|Gets the index of the row in its parent table.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| {            ignorePunct?: boolean            ignoreSpace?: boolean            matchCase?: boolean            matchPrefix?: boolean            matchSuffix?: boolean            matchWholeWord?: boolean            matchWildcards?: boolean        })](/javascript/api/word/word.tablerow#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|Performs a search with the specified SearchOptions on the scope of the row.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.tablerow#select-selectionmode-)|Selects the row and navigates the Word UI to it.|
||[setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.tablerow#setcellpadding-cellpaddinglocation--cellpadding-)|Sets cell padding in points.|
||[shadingColor](/javascript/api/word/word.tablerow#shadingcolor)|Gets and sets the shading color.|
||[values](/javascript/api/word/word.tablerow#values)|Gets and sets the text values in the row, as a 2D Javascript array.|
||[verticalAlignment](/javascript/api/word/word.tablerow#verticalalignment)|Gets and sets the vertical alignment of the cells in the row.|
|[TableRowCollection](/javascript/api/word/word.tablerowcollection)|[getFirst()](/javascript/api/word/word.tablerowcollection#getfirst--)|Gets the first row in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.tablerowcollection#getfirstornullobject--)|Gets the first row in this collection.|
||[items](/javascript/api/word/word.tablerowcollection#items)|Gets the loaded child items in this collection.|

## See also

- [Word JavaScript API Reference Documentation](/javascript/api/word)
- [Word JavaScript API requirement sets](word-api-requirement-sets.md)
