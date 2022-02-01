---
title: Word JavaScript API requirement set 1.3
description: 'Details about the WordApi 1.3 requirement set.'
ms.date: 03/09/2021
ms.prod: word
ms.localizationpriority: medium
---

# What's new in Word JavaScript API 1.3

WordApi 1.3 added more support for content controls and document-level settings.

## API list

The following table lists the APIs in Word JavaScript API requirement set 1.3. To view API reference documentation for all APIs supported by Word JavaScript API requirement set 1.3 or earlier, see [Word APIs in requirement set 1.3 or earlier](/javascript/api/word?view=word-js-1.3&preserve-view=true).

| Class | Fields | Description |
|:---|:---|:---|
|[Application](/javascript/api/word/word.application)|[createDocument(base64File?: string)](/javascript/api/word/word.application#word-word-application-createDocument-member(1))|Creates a new document by using an optional base64 encoded .docx file.|
|[Body](/javascript/api/word/word.body)|[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.body#word-word-body-getRange-member(1))|Gets the whole body, or the starting or ending point of the body, as a range.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.body#word-word-body-insertTable-member(1))|Inserts a table with the specified number of rows and columns.|
||[lists](/javascript/api/word/word.body#word-word-body-lists-member)|Gets the collection of list objects in the body.|
||[parentBody](/javascript/api/word/word.body#word-word-body-parentBody-member)|Gets the parent body of the body.|
||[parentBodyOrNullObject](/javascript/api/word/word.body#word-word-body-parentBodyOrNullObject-member)|Gets the parent body of the body.|
||[parentContentControlOrNullObject](/javascript/api/word/word.body#word-word-body-parentContentControlOrNullObject-member)|Gets the content control that contains the body.|
||[parentSection](/javascript/api/word/word.body#word-word-body-parentSection-member)|Gets the parent section of the body.|
||[parentSectionOrNullObject](/javascript/api/word/word.body#word-word-body-parentSectionOrNullObject-member)|Gets the parent section of the body.|
||[styleBuiltIn](/javascript/api/word/word.body#word-word-body-styleBuiltIn-member)|Gets or sets the built-in style name for the body.|
||[tables](/javascript/api/word/word.body#word-word-body-tables-member)|Gets the collection of table objects in the body.|
||[type](/javascript/api/word/word.body#word-word-body-type-member)|Gets the type of the body.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-getRange-member(1))|Gets the whole content control, or the starting or ending point of the content control, as a range.|
||[getTextRanges(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-getTextRanges-member(1))|Gets the text ranges in the content control by using punctuation marks and/or other ending marks.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-insertTable-member(1))|Inserts a table with the specified number of rows and columns into, or next to, a content control.|
||[lists](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-lists-member)|Gets the collection of list objects in the content control.|
||[parentBody](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parentBody-member)|Gets the parent body of the content control.|
||[parentContentControlOrNullObject](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parentContentControlOrNullObject-member)|Gets the content control that contains the content control.|
||[parentTable](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parentTable-member)|Gets the table that contains the content control.|
||[parentTableCell](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parentTableCell-member)|Gets the table cell that contains the content control.|
||[parentTableCellOrNullObject](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parentTableCellOrNullObject-member)|Gets the table cell that contains the content control.|
||[parentTableOrNullObject](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parentTableOrNullObject-member)|Gets the table that contains the content control.|
||[split(delimiters: string[], multiParagraphs?: boolean, trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-split-member(1))|Splits the content control into child ranges by using delimiters.|
||[styleBuiltIn](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-styleBuiltIn-member)|Gets or sets the built-in style name for the content control.|
||[subtype](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-subtype-member)|Gets the content control subtype.|
||[tables](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-tables-member)|Gets the collection of table objects in the content control.|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getByIdOrNullObject(id: number)](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getByIdOrNullObject-member(1))|Gets a content control by its identifier.|
||[getByTypes(types: Word.ContentControlType[])](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getByTypes-member(1))|Gets the content controls that have the specified types and/or subtypes.|
||[getFirst()](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getFirst-member(1))|Gets the first content control in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getFirstOrNullObject-member(1))|Gets the first content control in this collection.|
|[CustomProperty](/javascript/api/word/word.customproperty)|[delete()](/javascript/api/word/word.customproperty#word-word-customproperty-delete-member(1))|Deletes the custom property.|
||[key](/javascript/api/word/word.customproperty#word-word-customproperty-key-member)|Gets the key of the custom property.|
||[type](/javascript/api/word/word.customproperty#word-word-customproperty-type-member)|Gets the value type of the custom property.|
||[value](/javascript/api/word/word.customproperty#word-word-customproperty-value-member)|Gets or sets the value of the custom property.|
|[CustomPropertyCollection](/javascript/api/word/word.custompropertycollection)|[add(key: string, value: any)](/javascript/api/word/word.custompropertycollection#word-word-custompropertycollection-add-member(1))|Creates a new or sets an existing custom property.|
||[deleteAll()](/javascript/api/word/word.custompropertycollection#word-word-custompropertycollection-deleteAll-member(1))|Deletes all custom properties in this collection.|
||[getCount()](/javascript/api/word/word.custompropertycollection#word-word-custompropertycollection-getCount-member(1))|Gets the count of custom properties.|
||[getItem(key: string)](/javascript/api/word/word.custompropertycollection#word-word-custompropertycollection-getItem-member(1))|Gets a custom property object by its key, which is case-insensitive.|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.custompropertycollection#word-word-custompropertycollection-getItemOrNullObject-member(1))|Gets a custom property object by its key, which is case-insensitive.|
||[items](/javascript/api/word/word.custompropertycollection#word-word-custompropertycollection-items-member)|Gets the loaded child items in this collection.|
|[Document](/javascript/api/word/word.document)|[properties](/javascript/api/word/word.document#word-word-document-properties-member)|Gets the properties of the document.|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[body](/javascript/api/word/word.documentcreated#word-word-documentcreated-body-member)|Gets the body object of the document.|
||[contentControls](/javascript/api/word/word.documentcreated#word-word-documentcreated-contentControls-member)|Gets the collection of content control objects in the document.|
||[open()](/javascript/api/word/word.documentcreated#word-word-documentcreated-open-member(1))|Opens the document.|
||[properties](/javascript/api/word/word.documentcreated#word-word-documentcreated-properties-member)|Gets the properties of the document.|
||[save()](/javascript/api/word/word.documentcreated#word-word-documentcreated-save-member(1))|Saves the document.|
||[saved](/javascript/api/word/word.documentcreated#word-word-documentcreated-saved-member)|Indicates whether the changes in the document have been saved.|
||[sections](/javascript/api/word/word.documentcreated#word-word-documentcreated-sections-member)|Gets the collection of section objects in the document.|
|[DocumentProperties](/javascript/api/word/word.documentproperties)|[applicationName](/javascript/api/word/word.documentproperties#word-word-documentproperties-applicationName-member)|Gets the application name of the document.|
||[author](/javascript/api/word/word.documentproperties#word-word-documentproperties-author-member)|Gets or sets the author of the document.|
||[category](/javascript/api/word/word.documentproperties#word-word-documentproperties-category-member)|Gets or sets the category of the document.|
||[comments](/javascript/api/word/word.documentproperties#word-word-documentproperties-comments-member)|Gets or sets the comments of the document.|
||[company](/javascript/api/word/word.documentproperties#word-word-documentproperties-company-member)|Gets or sets the company of the document.|
||[creationDate](/javascript/api/word/word.documentproperties#word-word-documentproperties-creationDate-member)|Gets the creation date of the document.|
||[customProperties](/javascript/api/word/word.documentproperties#word-word-documentproperties-customProperties-member)|Gets the collection of custom properties of the document.|
||[format](/javascript/api/word/word.documentproperties#word-word-documentproperties-format-member)|Gets or sets the format of the document.|
||[keywords](/javascript/api/word/word.documentproperties#word-word-documentproperties-keywords-member)|Gets or sets the keywords of the document.|
||[lastAuthor](/javascript/api/word/word.documentproperties#word-word-documentproperties-lastAuthor-member)|Gets the last author of the document.|
||[lastPrintDate](/javascript/api/word/word.documentproperties#word-word-documentproperties-lastPrintDate-member)|Gets the last print date of the document.|
||[lastSaveTime](/javascript/api/word/word.documentproperties#word-word-documentproperties-lastSaveTime-member)|Gets the last save time of the document.|
||[manager](/javascript/api/word/word.documentproperties#word-word-documentproperties-manager-member)|Gets or sets the manager of the document.|
||[revisionNumber](/javascript/api/word/word.documentproperties#word-word-documentproperties-revisionNumber-member)|Gets the revision number of the document.|
||[security](/javascript/api/word/word.documentproperties#word-word-documentproperties-security-member)|Gets security settings of the document.|
||[subject](/javascript/api/word/word.documentproperties#word-word-documentproperties-subject-member)|Gets or sets the subject of the document.|
||[template](/javascript/api/word/word.documentproperties#word-word-documentproperties-template-member)|Gets the template of the document.|
||[title](/javascript/api/word/word.documentproperties#word-word-documentproperties-title-member)|Gets or sets the title of the document.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[getNext()](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-getNext-member(1))|Gets the next inline image.|
||[getNextOrNullObject()](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-getNextOrNullObject-member(1))|Gets the next inline image.|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-getRange-member(1))|Gets the picture, or the starting or ending point of the picture, as a range.|
||[parentContentControlOrNullObject](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-parentContentControlOrNullObject-member)|Gets the content control that contains the inline image.|
||[parentTable](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-parentTable-member)|Gets the table that contains the inline image.|
||[parentTableCell](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-parentTableCell-member)|Gets the table cell that contains the inline image.|
||[parentTableCellOrNullObject](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-parentTableCellOrNullObject-member)|Gets the table cell that contains the inline image.|
||[parentTableOrNullObject](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-parentTableOrNullObject-member)|Gets the table that contains the inline image.|
|[InlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|[getFirst()](/javascript/api/word/word.inlinepicturecollection#word-word-inlinepicturecollection-getFirst-member(1))|Gets the first inline image in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.inlinepicturecollection#word-word-inlinepicturecollection-getFirstOrNullObject-member(1))|Gets the first inline image in this collection.|
|[List](/javascript/api/word/word.list)|[getLevelParagraphs(level: number)](/javascript/api/word/word.list#word-word-list-getLevelParagraphs-member(1))|Gets the paragraphs that occur at the specified level in the list.|
||[getLevelString(level: number)](/javascript/api/word/word.list#word-word-list-getLevelString-member(1))|Gets the bullet, number, or picture at the specified level as a string.|
||[id](/javascript/api/word/word.list#word-word-list-id-member)|Gets the list's id.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.list#word-word-list-insertParagraph-member(1))|Inserts a paragraph at the specified location.|
||[levelExistences](/javascript/api/word/word.list#word-word-list-levelExistences-member)|Checks whether each of the 9 levels exists in the list.|
||[levelTypes](/javascript/api/word/word.list#word-word-list-levelTypes-member)|Gets all 9 level types in the list.|
||[paragraphs](/javascript/api/word/word.list#word-word-list-paragraphs-member)|Gets paragraphs in the list.|
||[setLevelAlignment(level: number, alignment: Word.Alignment)](/javascript/api/word/word.list#word-word-list-setLevelAlignment-member(1))|Sets the alignment of the bullet, number, or picture at the specified level in the list.|
||[setLevelBullet(level: number, listBullet: Word.ListBullet, charCode?: number, fontName?: string)](/javascript/api/word/word.list#word-word-list-setLevelBullet-member(1))|Sets the bullet format at the specified level in the list.|
||[setLevelIndents(level: number, textIndent: number, bulletNumberPictureIndent: number)](/javascript/api/word/word.list#word-word-list-setLevelIndents-member(1))|Sets the two indents of the specified level in the list.|
||[setLevelNumbering(level: number, listNumbering: Word.ListNumbering, formatString?: Array<string \| number>)](/javascript/api/word/word.list#word-word-list-setLevelNumbering-member(1))|Sets the numbering format at the specified level in the list.|
||[setLevelStartingNumber(level: number, startingNumber: number)](/javascript/api/word/word.list#word-word-list-setLevelStartingNumber-member(1))|Sets the starting number at the specified level in the list.|
|[ListCollection](/javascript/api/word/word.listcollection)|[getById(id: number)](/javascript/api/word/word.listcollection#word-word-listcollection-getById-member(1))|Gets a list by its identifier.|
||[getByIdOrNullObject(id: number)](/javascript/api/word/word.listcollection#word-word-listcollection-getByIdOrNullObject-member(1))|Gets a list by its identifier.|
||[getFirst()](/javascript/api/word/word.listcollection#word-word-listcollection-getFirst-member(1))|Gets the first list in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.listcollection#word-word-listcollection-getFirstOrNullObject-member(1))|Gets the first list in this collection.|
||[getItem(index: number)](/javascript/api/word/word.listcollection#word-word-listcollection-getItem-member(1))|Gets a list object by its index in the collection.|
||[items](/javascript/api/word/word.listcollection#word-word-listcollection-items-member)|Gets the loaded child items in this collection.|
|[ListItem](/javascript/api/word/word.listitem)|[getAncestor(parentOnly?: boolean)](/javascript/api/word/word.listitem#word-word-listitem-getAncestor-member(1))|Gets the list item parent, or the closest ancestor if the parent does not exist.|
||[getAncestorOrNullObject(parentOnly?: boolean)](/javascript/api/word/word.listitem#word-word-listitem-getAncestorOrNullObject-member(1))|Gets the list item parent, or the closest ancestor if the parent does not exist.|
||[getDescendants(directChildrenOnly?: boolean)](/javascript/api/word/word.listitem#word-word-listitem-getDescendants-member(1))|Gets all descendant list items of the list item.|
||[level](/javascript/api/word/word.listitem#word-word-listitem-level-member)|Gets or sets the level of the item in the list.|
||[listString](/javascript/api/word/word.listitem#word-word-listitem-listString-member)|Gets the list item bullet, number, or picture as a string.|
||[siblingIndex](/javascript/api/word/word.listitem#word-word-listitem-siblingIndex-member)|Gets the list item order number in relation to its siblings.|
|[Paragraph](/javascript/api/word/word.paragraph)|[attachToList(listId: number, level: number)](/javascript/api/word/word.paragraph#word-word-paragraph-attachToList-member(1))|Lets the paragraph join an existing list at the specified level.|
||[detachFromList()](/javascript/api/word/word.paragraph#word-word-paragraph-detachFromList-member(1))|Moves this paragraph out of its list, if the paragraph is a list item.|
||[getNext()](/javascript/api/word/word.paragraph#word-word-paragraph-getNext-member(1))|Gets the next paragraph.|
||[getNextOrNullObject()](/javascript/api/word/word.paragraph#word-word-paragraph-getNextOrNullObject-member(1))|Gets the next paragraph.|
||[getPrevious()](/javascript/api/word/word.paragraph#word-word-paragraph-getPrevious-member(1))|Gets the previous paragraph.|
||[getPreviousOrNullObject()](/javascript/api/word/word.paragraph#word-word-paragraph-getPreviousOrNullObject-member(1))|Gets the previous paragraph.|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.paragraph#word-word-paragraph-getRange-member(1))|Gets the whole paragraph, or the starting or ending point of the paragraph, as a range.|
||[getTextRanges(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.paragraph#word-word-paragraph-getTextRanges-member(1))|Gets the text ranges in the paragraph by using punctuation marks and/or other ending marks.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.paragraph#word-word-paragraph-insertTable-member(1))|Inserts a table with the specified number of rows and columns.|
||[isLastParagraph](/javascript/api/word/word.paragraph#word-word-paragraph-isLastParagraph-member)|Indicates the paragraph is the last one inside its parent body.|
||[isListItem](/javascript/api/word/word.paragraph#word-word-paragraph-isListItem-member)|Checks whether the paragraph is a list item.|
||[list](/javascript/api/word/word.paragraph#word-word-paragraph-list-member)|Gets the List to which this paragraph belongs.|
||[listItem](/javascript/api/word/word.paragraph#word-word-paragraph-listItem-member)|Gets the ListItem for the paragraph.|
||[listItemOrNullObject](/javascript/api/word/word.paragraph#word-word-paragraph-listItemOrNullObject-member)|Gets the ListItem for the paragraph.|
||[listOrNullObject](/javascript/api/word/word.paragraph#word-word-paragraph-listOrNullObject-member)|Gets the List to which this paragraph belongs.|
||[parentBody](/javascript/api/word/word.paragraph#word-word-paragraph-parentBody-member)|Gets the parent body of the paragraph.|
||[parentContentControlOrNullObject](/javascript/api/word/word.paragraph#word-word-paragraph-parentContentControlOrNullObject-member)|Gets the content control that contains the paragraph.|
||[parentTable](/javascript/api/word/word.paragraph#word-word-paragraph-parentTable-member)|Gets the table that contains the paragraph.|
||[parentTableCell](/javascript/api/word/word.paragraph#word-word-paragraph-parentTableCell-member)|Gets the table cell that contains the paragraph.|
||[parentTableCellOrNullObject](/javascript/api/word/word.paragraph#word-word-paragraph-parentTableCellOrNullObject-member)|Gets the table cell that contains the paragraph.|
||[parentTableOrNullObject](/javascript/api/word/word.paragraph#word-word-paragraph-parentTableOrNullObject-member)|Gets the table that contains the paragraph.|
||[split(delimiters: string[], trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.paragraph#word-word-paragraph-split-member(1))|Splits the paragraph into child ranges by using delimiters.|
||[startNewList()](/javascript/api/word/word.paragraph#word-word-paragraph-startNewList-member(1))|Starts a new list with this paragraph.|
||[styleBuiltIn](/javascript/api/word/word.paragraph#word-word-paragraph-styleBuiltIn-member)|Gets or sets the built-in style name for the paragraph.|
||[tableNestingLevel](/javascript/api/word/word.paragraph#word-word-paragraph-tableNestingLevel-member)|Gets the level of the paragraph's table.|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[getFirst()](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-getFirst-member(1))|Gets the first paragraph in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-getFirstOrNullObject-member(1))|Gets the first paragraph in this collection.|
||[getLast()](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-getLast-member(1))|Gets the last paragraph in this collection.|
||[getLastOrNullObject()](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-getLastOrNullObject-member(1))|Gets the last paragraph in this collection.|
|[Range](/javascript/api/word/word.range)|[compareLocationWith(range: Word.Range)](/javascript/api/word/word.range#word-word-range-compareLocationWith-member(1))|Compares this range's location with another range's location.|
||[expandTo(range: Word.Range)](/javascript/api/word/word.range#word-word-range-expandTo-member(1))|Returns a new range that extends from this range in either direction to cover another range.|
||[expandToOrNullObject(range: Word.Range)](/javascript/api/word/word.range#word-word-range-expandToOrNullObject-member(1))|Returns a new range that extends from this range in either direction to cover another range.|
||[getHyperlinkRanges()](/javascript/api/word/word.range#word-word-range-getHyperlinkRanges-member(1))|Gets hyperlink child ranges within the range.|
||[getNextTextRange(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#word-word-range-getNextTextRange-member(1))|Gets the next text range by using punctuation marks and/or other ending marks.|
||[getNextTextRangeOrNullObject(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#word-word-range-getNextTextRangeOrNullObject-member(1))|Gets the next text range by using punctuation marks and/or other ending marks.|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.range#word-word-range-getRange-member(1))|Clones the range, or gets the starting or ending point of the range as a new range.|
||[getTextRanges(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#word-word-range-getTextRanges-member(1))|Gets the text child ranges in the range by using punctuation marks and/or other ending marks.|
||[hyperlink](/javascript/api/word/word.range#word-word-range-hyperlink-member)|Gets the first hyperlink in the range, or sets a hyperlink on the range.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.range#word-word-range-insertTable-member(1))|Inserts a table with the specified number of rows and columns.|
||[intersectWith(range: Word.Range)](/javascript/api/word/word.range#word-word-range-intersectWith-member(1))|Returns a new range as the intersection of this range with another range.|
||[intersectWithOrNullObject(range: Word.Range)](/javascript/api/word/word.range#word-word-range-intersectWithOrNullObject-member(1))|Returns a new range as the intersection of this range with another range.|
||[isEmpty](/javascript/api/word/word.range#word-word-range-isEmpty-member)|Checks whether the range length is zero.|
||[lists](/javascript/api/word/word.range#word-word-range-lists-member)|Gets the collection of list objects in the range.|
||[parentBody](/javascript/api/word/word.range#word-word-range-parentBody-member)|Gets the parent body of the range.|
||[parentContentControlOrNullObject](/javascript/api/word/word.range#word-word-range-parentContentControlOrNullObject-member)|Gets the content control that contains the range.|
||[parentTable](/javascript/api/word/word.range#word-word-range-parentTable-member)|Gets the table that contains the range.|
||[parentTableCell](/javascript/api/word/word.range#word-word-range-parentTableCell-member)|Gets the table cell that contains the range.|
||[parentTableCellOrNullObject](/javascript/api/word/word.range#word-word-range-parentTableCellOrNullObject-member)|Gets the table cell that contains the range.|
||[parentTableOrNullObject](/javascript/api/word/word.range#word-word-range-parentTableOrNullObject-member)|Gets the table that contains the range.|
||[split(delimiters: string[], multiParagraphs?: boolean, trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.range#word-word-range-split-member(1))|Splits the range into child ranges by using delimiters.|
||[styleBuiltIn](/javascript/api/word/word.range#word-word-range-styleBuiltIn-member)|Gets or sets the built-in style name for the range.|
||[tables](/javascript/api/word/word.range#word-word-range-tables-member)|Gets the collection of table objects in the range.|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[getFirst()](/javascript/api/word/word.rangecollection#word-word-rangecollection-getFirst-member(1))|Gets the first range in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.rangecollection#word-word-rangecollection-getFirstOrNullObject-member(1))|Gets the first range in this collection.|
|[RequestContext](/javascript/api/word/word.requestcontext)|[application](/javascript/api/word/word.requestcontext#word-word-requestcontext-application-member)|[Api set: WordApi 1.3] *|
|[Section](/javascript/api/word/word.section)|[getNext()](/javascript/api/word/word.section#word-word-section-getNext-member(1))|Gets the next section.|
||[getNextOrNullObject()](/javascript/api/word/word.section#word-word-section-getNextOrNullObject-member(1))|Gets the next section.|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[getFirst()](/javascript/api/word/word.sectioncollection#word-word-sectioncollection-getFirst-member(1))|Gets the first section in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.sectioncollection#word-word-sectioncollection-getFirstOrNullObject-member(1))|Gets the first section in this collection.|
|[Table](/javascript/api/word/word.table)|[addColumns(insertLocation: Word.InsertLocation, columnCount: number, values?: string[][])](/javascript/api/word/word.table#word-word-table-addColumns-member(1))|Adds columns to the start or end of the table, using the first or last existing column as a template.|
||[addRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[][])](/javascript/api/word/word.table#word-word-table-addRows-member(1))|Adds rows to the start or end of the table, using the first or last existing row as a template.|
||[alignment](/javascript/api/word/word.table#word-word-table-alignment-member)|Gets or sets the alignment of the table against the page column.|
||[autoFitWindow()](/javascript/api/word/word.table#word-word-table-autoFitWindow-member(1))|Autofits the table columns to the width of the window.|
||[clear()](/javascript/api/word/word.table#word-word-table-clear-member(1))|Clears the contents of the table.|
||[delete()](/javascript/api/word/word.table#word-word-table-delete-member(1))|Deletes the entire table.|
||[deleteColumns(columnIndex: number, columnCount?: number)](/javascript/api/word/word.table#word-word-table-deleteColumns-member(1))|Deletes specific columns.|
||[deleteRows(rowIndex: number, rowCount?: number)](/javascript/api/word/word.table#word-word-table-deleteRows-member(1))|Deletes specific rows.|
||[distributeColumns()](/javascript/api/word/word.table#word-word-table-distributeColumns-member(1))|Distributes the column widths evenly.|
||[font](/javascript/api/word/word.table#word-word-table-font-member)|Gets the font.|
||[getBorder(borderLocation: Word.BorderLocation)](/javascript/api/word/word.table#word-word-table-getBorder-member(1))|Gets the border style for the specified border.|
||[getCell(rowIndex: number, cellIndex: number)](/javascript/api/word/word.table#word-word-table-getCell-member(1))|Gets the table cell at a specified row and column.|
||[getCellOrNullObject(rowIndex: number, cellIndex: number)](/javascript/api/word/word.table#word-word-table-getCellOrNullObject-member(1))|Gets the table cell at a specified row and column.|
||[getCellPadding(cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.table#word-word-table-getCellPadding-member(1))|Gets cell padding in points.|
||[getNext()](/javascript/api/word/word.table#word-word-table-getNext-member(1))|Gets the next table.|
||[getNextOrNullObject()](/javascript/api/word/word.table#word-word-table-getNextOrNullObject-member(1))|Gets the next table.|
||[getParagraphAfter()](/javascript/api/word/word.table#word-word-table-getParagraphAfter-member(1))|Gets the paragraph after the table.|
||[getParagraphAfterOrNullObject()](/javascript/api/word/word.table#word-word-table-getParagraphAfterOrNullObject-member(1))|Gets the paragraph after the table.|
||[getParagraphBefore()](/javascript/api/word/word.table#word-word-table-getParagraphBefore-member(1))|Gets the paragraph before the table.|
||[getParagraphBeforeOrNullObject()](/javascript/api/word/word.table#word-word-table-getParagraphBeforeOrNullObject-member(1))|Gets the paragraph before the table.|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.table#word-word-table-getRange-member(1))|Gets the range that contains this table, or the range at the start or end of the table.|
||[headerRowCount](/javascript/api/word/word.table#word-word-table-headerRowCount-member)|Gets and sets the number of header rows.|
||[horizontalAlignment](/javascript/api/word/word.table#word-word-table-horizontalAlignment-member)|Gets and sets the horizontal alignment of every cell in the table.|
||[ignorePunct](/javascript/api/word/word.table#word-word-table-ignorePunct-member)||
||[ignoreSpace](/javascript/api/word/word.table#word-word-table-ignoreSpace-member)||
||[insertContentControl()](/javascript/api/word/word.table#word-word-table-insertContentControl-member(1))|Inserts a content control on the table.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.table#word-word-table-insertParagraph-member(1))|Inserts a paragraph at the specified location.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.table#word-word-table-insertTable-member(1))|Inserts a table with the specified number of rows and columns.|
||[isUniform](/javascript/api/word/word.table#word-word-table-isUniform-member)|Indicates whether all of the table rows are uniform.|
||[matchCase](/javascript/api/word/word.table#word-word-table-matchCase-member)||
||[matchPrefix](/javascript/api/word/word.table#word-word-table-matchPrefix-member)||
||[matchSuffix](/javascript/api/word/word.table#word-word-table-matchSuffix-member)||
||[matchWholeWord](/javascript/api/word/word.table#word-word-table-matchWholeWord-member)||
||[matchWildcards](/javascript/api/word/word.table#word-word-table-matchWildcards-member)||
||[nestingLevel](/javascript/api/word/word.table#word-word-table-nestingLevel-member)|Gets the nesting level of the table.|
||[parentBody](/javascript/api/word/word.table#word-word-table-parentBody-member)|Gets the parent body of the table.|
||[parentContentControl](/javascript/api/word/word.table#word-word-table-parentContentControl-member)|Gets the content control that contains the table.|
||[parentContentControlOrNullObject](/javascript/api/word/word.table#word-word-table-parentContentControlOrNullObject-member)|Gets the content control that contains the table.|
||[parentTable](/javascript/api/word/word.table#word-word-table-parentTable-member)|Gets the table that contains this table.|
||[parentTableCell](/javascript/api/word/word.table#word-word-table-parentTableCell-member)|Gets the table cell that contains this table.|
||[parentTableCellOrNullObject](/javascript/api/word/word.table#word-word-table-parentTableCellOrNullObject-member)|Gets the table cell that contains this table.|
||[parentTableOrNullObject](/javascript/api/word/word.table#word-word-table-parentTableOrNullObject-member)|Gets the table that contains this table.|
||[rowCount](/javascript/api/word/word.table#word-word-table-rowCount-member)|Gets the number of rows in the table.|
||[rows](/javascript/api/word/word.table#word-word-table-rows-member)|Gets all of the table rows.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| {            ignorePunct?: boolean            ignoreSpace?: boolean            matchCase?: boolean            matchPrefix?: boolean            matchSuffix?: boolean            matchWholeWord?: boolean            matchWildcards?: boolean        })](/javascript/api/word/word.table#word-word-table-search-member(1))|Performs a search with the specified SearchOptions on the scope of the table object.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.table#word-word-table-select-member(1))|Selects the table, or the position at the start or end of the table, and navigates the Word UI to it.|
||[setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.table#word-word-table-setCellPadding-member(1))|Sets cell padding in points.|
||[shadingColor](/javascript/api/word/word.table#word-word-table-shadingColor-member)|Gets and sets the shading color.|
||[style](/javascript/api/word/word.table#word-word-table-style-member)|Gets or sets the style name for the table.|
||[styleBandedColumns](/javascript/api/word/word.table#word-word-table-styleBandedColumns-member)|Gets and sets whether the table has banded columns.|
||[styleBandedRows](/javascript/api/word/word.table#word-word-table-styleBandedRows-member)|Gets and sets whether the table has banded rows.|
||[styleBuiltIn](/javascript/api/word/word.table#word-word-table-styleBuiltIn-member)|Gets or sets the built-in style name for the table.|
||[styleFirstColumn](/javascript/api/word/word.table#word-word-table-styleFirstColumn-member)|Gets and sets whether the table has a first column with a special style.|
||[styleLastColumn](/javascript/api/word/word.table#word-word-table-styleLastColumn-member)|Gets and sets whether the table has a last column with a special style.|
||[styleTotalRow](/javascript/api/word/word.table#word-word-table-styleTotalRow-member)|Gets and sets whether the table has a total (last) row with a special style.|
||[tables](/javascript/api/word/word.table#word-word-table-tables-member)|Gets the child tables nested one level deeper.|
||[values](/javascript/api/word/word.table#word-word-table-values-member)|Gets and sets the text values in the table, as a 2D Javascript array.|
||[verticalAlignment](/javascript/api/word/word.table#word-word-table-verticalAlignment-member)|Gets and sets the vertical alignment of every cell in the table.|
||[width](/javascript/api/word/word.table#word-word-table-width-member)|Gets and sets the width of the table in points.|
|[TableBorder](/javascript/api/word/word.tableborder)|[color](/javascript/api/word/word.tableborder#word-word-tableborder-color-member)|Gets or sets the table border color.|
||[type](/javascript/api/word/word.tableborder#word-word-tableborder-type-member)|Gets or sets the type of the table border.|
||[width](/javascript/api/word/word.tableborder#word-word-tableborder-width-member)|Gets or sets the width, in points, of the table border.|
|[TableCell](/javascript/api/word/word.tablecell)|[body](/javascript/api/word/word.tablecell#word-word-tablecell-body-member)|Gets the body object of the cell.|
||[cellIndex](/javascript/api/word/word.tablecell#word-word-tablecell-cellIndex-member)|Gets the index of the cell in its row.|
||[columnWidth](/javascript/api/word/word.tablecell#word-word-tablecell-columnWidth-member)|Gets and sets the width of the cell's column in points.|
||[deleteColumn()](/javascript/api/word/word.tablecell#word-word-tablecell-deleteColumn-member(1))|Deletes the column containing this cell.|
||[deleteRow()](/javascript/api/word/word.tablecell#word-word-tablecell-deleteRow-member(1))|Deletes the row containing this cell.|
||[getBorder(borderLocation: Word.BorderLocation)](/javascript/api/word/word.tablecell#word-word-tablecell-getBorder-member(1))|Gets the border style for the specified border.|
||[getCellPadding(cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.tablecell#word-word-tablecell-getCellPadding-member(1))|Gets cell padding in points.|
||[getNext()](/javascript/api/word/word.tablecell#word-word-tablecell-getNext-member(1))|Gets the next cell.|
||[getNextOrNullObject()](/javascript/api/word/word.tablecell#word-word-tablecell-getNextOrNullObject-member(1))|Gets the next cell.|
||[horizontalAlignment](/javascript/api/word/word.tablecell#word-word-tablecell-horizontalAlignment-member)|Gets and sets the horizontal alignment of the cell.|
||[insertColumns(insertLocation: Word.InsertLocation, columnCount: number, values?: string[][])](/javascript/api/word/word.tablecell#word-word-tablecell-insertColumns-member(1))|Adds columns to the left or right of the cell, using the cell's column as a template.|
||[insertRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[][])](/javascript/api/word/word.tablecell#word-word-tablecell-insertRows-member(1))|Inserts rows above or below the cell, using the cell's row as a template.|
||[parentRow](/javascript/api/word/word.tablecell#word-word-tablecell-parentRow-member)|Gets the parent row of the cell.|
||[parentTable](/javascript/api/word/word.tablecell#word-word-tablecell-parentTable-member)|Gets the parent table of the cell.|
||[rowIndex](/javascript/api/word/word.tablecell#word-word-tablecell-rowIndex-member)|Gets the index of the cell's row in the table.|
||[setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.tablecell#word-word-tablecell-setCellPadding-member(1))|Sets cell padding in points.|
||[shadingColor](/javascript/api/word/word.tablecell#word-word-tablecell-shadingColor-member)|Gets or sets the shading color of the cell.|
||[value](/javascript/api/word/word.tablecell#word-word-tablecell-value-member)|Gets and sets the text of the cell.|
||[verticalAlignment](/javascript/api/word/word.tablecell#word-word-tablecell-verticalAlignment-member)|Gets and sets the vertical alignment of the cell.|
||[width](/javascript/api/word/word.tablecell#word-word-tablecell-width-member)|Gets the width of the cell in points.|
|[TableCellCollection](/javascript/api/word/word.tablecellcollection)|[getFirst()](/javascript/api/word/word.tablecellcollection#word-word-tablecellcollection-getFirst-member(1))|Gets the first table cell in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.tablecellcollection#word-word-tablecellcollection-getFirstOrNullObject-member(1))|Gets the first table cell in this collection.|
||[items](/javascript/api/word/word.tablecellcollection#word-word-tablecellcollection-items-member)|Gets the loaded child items in this collection.|
|[TableCollection](/javascript/api/word/word.tablecollection)|[getFirst()](/javascript/api/word/word.tablecollection#word-word-tablecollection-getFirst-member(1))|Gets the first table in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.tablecollection#word-word-tablecollection-getFirstOrNullObject-member(1))|Gets the first table in this collection.|
||[items](/javascript/api/word/word.tablecollection#word-word-tablecollection-items-member)|Gets the loaded child items in this collection.|
|[TableRow](/javascript/api/word/word.tablerow)|[cellCount](/javascript/api/word/word.tablerow#word-word-tablerow-cellCount-member)|Gets the number of cells in the row.|
||[cells](/javascript/api/word/word.tablerow#word-word-tablerow-cells-member)|Gets cells.|
||[clear()](/javascript/api/word/word.tablerow#word-word-tablerow-clear-member(1))|Clears the contents of the row.|
||[delete()](/javascript/api/word/word.tablerow#word-word-tablerow-delete-member(1))|Deletes the entire row.|
||[font](/javascript/api/word/word.tablerow#word-word-tablerow-font-member)|Gets the font.|
||[getBorder(borderLocation: Word.BorderLocation)](/javascript/api/word/word.tablerow#word-word-tablerow-getBorder-member(1))|Gets the border style of the cells in the row.|
||[getCellPadding(cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.tablerow#word-word-tablerow-getCellPadding-member(1))|Gets cell padding in points.|
||[getNext()](/javascript/api/word/word.tablerow#word-word-tablerow-getNext-member(1))|Gets the next row.|
||[getNextOrNullObject()](/javascript/api/word/word.tablerow#word-word-tablerow-getNextOrNullObject-member(1))|Gets the next row.|
||[horizontalAlignment](/javascript/api/word/word.tablerow#word-word-tablerow-horizontalAlignment-member)|Gets and sets the horizontal alignment of every cell in the row.|
||[ignorePunct](/javascript/api/word/word.tablerow#word-word-tablerow-ignorePunct-member)||
||[ignoreSpace](/javascript/api/word/word.tablerow#word-word-tablerow-ignoreSpace-member)||
||[insertRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[][])](/javascript/api/word/word.tablerow#word-word-tablerow-insertRows-member(1))|Inserts rows using this row as a template.|
||[isHeader](/javascript/api/word/word.tablerow#word-word-tablerow-isHeader-member)|Checks whether the row is a header row.|
||[matchCase](/javascript/api/word/word.tablerow#word-word-tablerow-matchCase-member)||
||[matchPrefix](/javascript/api/word/word.tablerow#word-word-tablerow-matchPrefix-member)||
||[matchSuffix](/javascript/api/word/word.tablerow#word-word-tablerow-matchSuffix-member)||
||[matchWholeWord](/javascript/api/word/word.tablerow#word-word-tablerow-matchWholeWord-member)||
||[matchWildcards](/javascript/api/word/word.tablerow#word-word-tablerow-matchWildcards-member)||
||[parentTable](/javascript/api/word/word.tablerow#word-word-tablerow-parentTable-member)|Gets parent table.|
||[preferredHeight](/javascript/api/word/word.tablerow#word-word-tablerow-preferredHeight-member)|Gets and sets the preferred height of the row in points.|
||[rowIndex](/javascript/api/word/word.tablerow#word-word-tablerow-rowIndex-member)|Gets the index of the row in its parent table.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| {            ignorePunct?: boolean            ignoreSpace?: boolean            matchCase?: boolean            matchPrefix?: boolean            matchSuffix?: boolean            matchWholeWord?: boolean            matchWildcards?: boolean        })](/javascript/api/word/word.tablerow#word-word-tablerow-search-member(1))|Performs a search with the specified SearchOptions on the scope of the row.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.tablerow#word-word-tablerow-select-member(1))|Selects the row and navigates the Word UI to it.|
||[setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.tablerow#word-word-tablerow-setCellPadding-member(1))|Sets cell padding in points.|
||[shadingColor](/javascript/api/word/word.tablerow#word-word-tablerow-shadingColor-member)|Gets and sets the shading color.|
||[values](/javascript/api/word/word.tablerow#word-word-tablerow-values-member)|Gets and sets the text values in the row, as a 2D Javascript array.|
||[verticalAlignment](/javascript/api/word/word.tablerow#word-word-tablerow-verticalAlignment-member)|Gets and sets the vertical alignment of the cells in the row.|
|[TableRowCollection](/javascript/api/word/word.tablerowcollection)|[getFirst()](/javascript/api/word/word.tablerowcollection#word-word-tablerowcollection-getFirst-member(1))|Gets the first row in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.tablerowcollection#word-word-tablerowcollection-getFirstOrNullObject-member(1))|Gets the first row in this collection.|
||[items](/javascript/api/word/word.tablerowcollection#word-word-tablerowcollection-items-member)|Gets the loaded child items in this collection.|

## See also

- [Word JavaScript API Reference Documentation](/javascript/api/word)
- [Word JavaScript API requirement sets](word-api-requirement-sets.md)
