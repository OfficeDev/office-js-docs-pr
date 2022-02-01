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
|[Application](/javascript/api/word/word.application)|[createDocument(base64File?: string)](/javascript/api/word/word.application#word-word-application-createdocument-member(1))|Creates a new document by using an optional base64 encoded .docx file.|
|[Body](/javascript/api/word/word.body)|[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.body#word-word-body-getrange-member(1))|Gets the whole body, or the starting or ending point of the body, as a range.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.body#word-word-body-inserttable-member(1))|Inserts a table with the specified number of rows and columns.|
||[lists](/javascript/api/word/word.body#word-word-body-lists-member)|Gets the collection of list objects in the body.|
||[parentBody](/javascript/api/word/word.body#word-word-body-parentbody-member)|Gets the parent body of the body.|
||[parentBodyOrNullObject](/javascript/api/word/word.body#word-word-body-parentbodyornullobject-member)|Gets the parent body of the body.|
||[parentContentControlOrNullObject](/javascript/api/word/word.body#word-word-body-parentcontentcontrolornullobject-member)|Gets the content control that contains the body.|
||[parentSection](/javascript/api/word/word.body#word-word-body-parentsection-member)|Gets the parent section of the body.|
||[parentSectionOrNullObject](/javascript/api/word/word.body#word-word-body-parentsectionornullobject-member)|Gets the parent section of the body.|
||[styleBuiltIn](/javascript/api/word/word.body#word-word-body-stylebuiltin-member)|Gets or sets the built-in style name for the body.|
||[tables](/javascript/api/word/word.body#word-word-body-tables-member)|Gets the collection of table objects in the body.|
||[type](/javascript/api/word/word.body#word-word-body-type-member)|Gets the type of the body.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-getrange-member(1))|Gets the whole content control, or the starting or ending point of the content control, as a range.|
||[getTextRanges(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-gettextranges-member(1))|Gets the text ranges in the content control by using punctuation marks and/or other ending marks.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-inserttable-member(1))|Inserts a table with the specified number of rows and columns into, or next to, a content control.|
||[lists](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-lists-member)|Gets the collection of list objects in the content control.|
||[parentBody](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parentbody-member)|Gets the parent body of the content control.|
||[parentContentControlOrNullObject](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parentcontentcontrolornullobject-member)|Gets the content control that contains the content control.|
||[parentTable](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parenttable-member)|Gets the table that contains the content control.|
||[parentTableCell](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parenttablecell-member)|Gets the table cell that contains the content control.|
||[parentTableCellOrNullObject](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parenttablecellornullobject-member)|Gets the table cell that contains the content control.|
||[parentTableOrNullObject](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parenttableornullobject-member)|Gets the table that contains the content control.|
||[split(delimiters: string[], multiParagraphs?: boolean, trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-split-member(1))|Splits the content control into child ranges by using delimiters.|
||[styleBuiltIn](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-stylebuiltin-member)|Gets or sets the built-in style name for the content control.|
||[subtype](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-subtype-member)|Gets the content control subtype.|
||[tables](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-tables-member)|Gets the collection of table objects in the content control.|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getByIdOrNullObject(id: number)](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getbyidornullobject-member(1))|Gets a content control by its identifier.|
||[getByTypes(types: Word.ContentControlType[])](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getbytypes-member(1))|Gets the content controls that have the specified types and/or subtypes.|
||[getFirst()](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getfirst-member(1))|Gets the first content control in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getfirstornullobject-member(1))|Gets the first content control in this collection.|
|[CustomProperty](/javascript/api/word/word.customproperty)|[delete()](/javascript/api/word/word.customproperty#word-word-customproperty-delete-member(1))|Deletes the custom property.|
||[key](/javascript/api/word/word.customproperty#word-word-customproperty-key-member)|Gets the key of the custom property.|
||[type](/javascript/api/word/word.customproperty#word-word-customproperty-type-member)|Gets the value type of the custom property.|
||[value](/javascript/api/word/word.customproperty#word-word-customproperty-value-member)|Gets or sets the value of the custom property.|
|[CustomPropertyCollection](/javascript/api/word/word.custompropertycollection)|[add(key: string, value: any)](/javascript/api/word/word.custompropertycollection#word-word-custompropertycollection-add-member(1))|Creates a new or sets an existing custom property.|
||[deleteAll()](/javascript/api/word/word.custompropertycollection#word-word-custompropertycollection-deleteall-member(1))|Deletes all custom properties in this collection.|
||[getCount()](/javascript/api/word/word.custompropertycollection#word-word-custompropertycollection-getcount-member(1))|Gets the count of custom properties.|
||[getItem(key: string)](/javascript/api/word/word.custompropertycollection#word-word-custompropertycollection-getitem-member(1))|Gets a custom property object by its key, which is case-insensitive.|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.custompropertycollection#word-word-custompropertycollection-getitemornullobject-member(1))|Gets a custom property object by its key, which is case-insensitive.|
||[items](/javascript/api/word/word.custompropertycollection#word-word-custompropertycollection-items-member)|Gets the loaded child items in this collection.|
|[Document](/javascript/api/word/word.document)|[properties](/javascript/api/word/word.document#word-word-document-properties-member)|Gets the properties of the document.|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[body](/javascript/api/word/word.documentcreated#word-word-documentcreated-body-member)|Gets the body object of the document.|
||[contentControls](/javascript/api/word/word.documentcreated#word-word-documentcreated-contentcontrols-member)|Gets the collection of content control objects in the document.|
||[open()](/javascript/api/word/word.documentcreated#word-word-documentcreated-open-member(1))|Opens the document.|
||[properties](/javascript/api/word/word.documentcreated#word-word-documentcreated-properties-member)|Gets the properties of the document.|
||[save()](/javascript/api/word/word.documentcreated#word-word-documentcreated-save-member(1))|Saves the document.|
||[saved](/javascript/api/word/word.documentcreated#word-word-documentcreated-saved-member)|Indicates whether the changes in the document have been saved.|
||[sections](/javascript/api/word/word.documentcreated#word-word-documentcreated-sections-member)|Gets the collection of section objects in the document.|
|[DocumentProperties](/javascript/api/word/word.documentproperties)|[applicationName](/javascript/api/word/word.documentproperties#word-word-documentproperties-applicationname-member)|Gets the application name of the document.|
||[author](/javascript/api/word/word.documentproperties#word-word-documentproperties-author-member)|Gets or sets the author of the document.|
||[category](/javascript/api/word/word.documentproperties#word-word-documentproperties-category-member)|Gets or sets the category of the document.|
||[comments](/javascript/api/word/word.documentproperties#word-word-documentproperties-comments-member)|Gets or sets the comments of the document.|
||[company](/javascript/api/word/word.documentproperties#word-word-documentproperties-company-member)|Gets or sets the company of the document.|
||[creationDate](/javascript/api/word/word.documentproperties#word-word-documentproperties-creationdate-member)|Gets the creation date of the document.|
||[customProperties](/javascript/api/word/word.documentproperties#word-word-documentproperties-customproperties-member)|Gets the collection of custom properties of the document.|
||[format](/javascript/api/word/word.documentproperties#word-word-documentproperties-format-member)|Gets or sets the format of the document.|
||[keywords](/javascript/api/word/word.documentproperties#word-word-documentproperties-keywords-member)|Gets or sets the keywords of the document.|
||[lastAuthor](/javascript/api/word/word.documentproperties#word-word-documentproperties-lastauthor-member)|Gets the last author of the document.|
||[lastPrintDate](/javascript/api/word/word.documentproperties#word-word-documentproperties-lastprintdate-member)|Gets the last print date of the document.|
||[lastSaveTime](/javascript/api/word/word.documentproperties#word-word-documentproperties-lastsavetime-member)|Gets the last save time of the document.|
||[manager](/javascript/api/word/word.documentproperties#word-word-documentproperties-manager-member)|Gets or sets the manager of the document.|
||[revisionNumber](/javascript/api/word/word.documentproperties#word-word-documentproperties-revisionnumber-member)|Gets the revision number of the document.|
||[security](/javascript/api/word/word.documentproperties#word-word-documentproperties-security-member)|Gets security settings of the document.|
||[subject](/javascript/api/word/word.documentproperties#word-word-documentproperties-subject-member)|Gets or sets the subject of the document.|
||[template](/javascript/api/word/word.documentproperties#word-word-documentproperties-template-member)|Gets the template of the document.|
||[title](/javascript/api/word/word.documentproperties#word-word-documentproperties-title-member)|Gets or sets the title of the document.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[getNext()](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-getnext-member(1))|Gets the next inline image.|
||[getNextOrNullObject()](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-getnextornullobject-member(1))|Gets the next inline image.|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-getrange-member(1))|Gets the picture, or the starting or ending point of the picture, as a range.|
||[parentContentControlOrNullObject](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-parentcontentcontrolornullobject-member)|Gets the content control that contains the inline image.|
||[parentTable](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-parenttable-member)|Gets the table that contains the inline image.|
||[parentTableCell](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-parenttablecell-member)|Gets the table cell that contains the inline image.|
||[parentTableCellOrNullObject](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-parenttablecellornullobject-member)|Gets the table cell that contains the inline image.|
||[parentTableOrNullObject](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-parenttableornullobject-member)|Gets the table that contains the inline image.|
|[InlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|[getFirst()](/javascript/api/word/word.inlinepicturecollection#word-word-inlinepicturecollection-getfirst-member(1))|Gets the first inline image in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.inlinepicturecollection#word-word-inlinepicturecollection-getfirstornullobject-member(1))|Gets the first inline image in this collection.|
|[List](/javascript/api/word/word.list)|[getLevelParagraphs(level: number)](/javascript/api/word/word.list#word-word-list-getlevelparagraphs-member(1))|Gets the paragraphs that occur at the specified level in the list.|
||[getLevelString(level: number)](/javascript/api/word/word.list#word-word-list-getlevelstring-member(1))|Gets the bullet, number, or picture at the specified level as a string.|
||[id](/javascript/api/word/word.list#word-word-list-id-member)|Gets the list's id.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.list#word-word-list-insertparagraph-member(1))|Inserts a paragraph at the specified location.|
||[levelExistences](/javascript/api/word/word.list#word-word-list-levelexistences-member)|Checks whether each of the 9 levels exists in the list.|
||[levelTypes](/javascript/api/word/word.list#word-word-list-leveltypes-member)|Gets all 9 level types in the list.|
||[paragraphs](/javascript/api/word/word.list#word-word-list-paragraphs-member)|Gets paragraphs in the list.|
||[setLevelAlignment(level: number, alignment: Word.Alignment)](/javascript/api/word/word.list#word-word-list-setlevelalignment-member(1))|Sets the alignment of the bullet, number, or picture at the specified level in the list.|
||[setLevelBullet(level: number, listBullet: Word.ListBullet, charCode?: number, fontName?: string)](/javascript/api/word/word.list#word-word-list-setlevelbullet-member(1))|Sets the bullet format at the specified level in the list.|
||[setLevelIndents(level: number, textIndent: number, bulletNumberPictureIndent: number)](/javascript/api/word/word.list#word-word-list-setlevelindents-member(1))|Sets the two indents of the specified level in the list.|
||[setLevelNumbering(level: number, listNumbering: Word.ListNumbering, formatString?: Array<string \| number>)](/javascript/api/word/word.list#word-word-list-setlevelnumbering-member(1))|Sets the numbering format at the specified level in the list.|
||[setLevelStartingNumber(level: number, startingNumber: number)](/javascript/api/word/word.list#word-word-list-setlevelstartingnumber-member(1))|Sets the starting number at the specified level in the list.|
|[ListCollection](/javascript/api/word/word.listcollection)|[getById(id: number)](/javascript/api/word/word.listcollection#word-word-listcollection-getbyid-member(1))|Gets a list by its identifier.|
||[getByIdOrNullObject(id: number)](/javascript/api/word/word.listcollection#word-word-listcollection-getbyidornullobject-member(1))|Gets a list by its identifier.|
||[getFirst()](/javascript/api/word/word.listcollection#word-word-listcollection-getfirst-member(1))|Gets the first list in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.listcollection#word-word-listcollection-getfirstornullobject-member(1))|Gets the first list in this collection.|
||[getItem(index: number)](/javascript/api/word/word.listcollection#word-word-listcollection-getitem-member(1))|Gets a list object by its index in the collection.|
||[items](/javascript/api/word/word.listcollection#word-word-listcollection-items-member)|Gets the loaded child items in this collection.|
|[ListItem](/javascript/api/word/word.listitem)|[getAncestor(parentOnly?: boolean)](/javascript/api/word/word.listitem#word-word-listitem-getancestor-member(1))|Gets the list item parent, or the closest ancestor if the parent does not exist.|
||[getAncestorOrNullObject(parentOnly?: boolean)](/javascript/api/word/word.listitem#word-word-listitem-getancestorornullobject-member(1))|Gets the list item parent, or the closest ancestor if the parent does not exist.|
||[getDescendants(directChildrenOnly?: boolean)](/javascript/api/word/word.listitem#word-word-listitem-getdescendants-member(1))|Gets all descendant list items of the list item.|
||[level](/javascript/api/word/word.listitem#word-word-listitem-level-member)|Gets or sets the level of the item in the list.|
||[listString](/javascript/api/word/word.listitem#word-word-listitem-liststring-member)|Gets the list item bullet, number, or picture as a string.|
||[siblingIndex](/javascript/api/word/word.listitem#word-word-listitem-siblingindex-member)|Gets the list item order number in relation to its siblings.|
|[Paragraph](/javascript/api/word/word.paragraph)|[attachToList(listId: number, level: number)](/javascript/api/word/word.paragraph#word-word-paragraph-attachtolist-member(1))|Lets the paragraph join an existing list at the specified level.|
||[detachFromList()](/javascript/api/word/word.paragraph#word-word-paragraph-detachfromlist-member(1))|Moves this paragraph out of its list, if the paragraph is a list item.|
||[getNext()](/javascript/api/word/word.paragraph#word-word-paragraph-getnext-member(1))|Gets the next paragraph.|
||[getNextOrNullObject()](/javascript/api/word/word.paragraph#word-word-paragraph-getnextornullobject-member(1))|Gets the next paragraph.|
||[getPrevious()](/javascript/api/word/word.paragraph#word-word-paragraph-getprevious-member(1))|Gets the previous paragraph.|
||[getPreviousOrNullObject()](/javascript/api/word/word.paragraph#word-word-paragraph-getpreviousornullobject-member(1))|Gets the previous paragraph.|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.paragraph#word-word-paragraph-getrange-member(1))|Gets the whole paragraph, or the starting or ending point of the paragraph, as a range.|
||[getTextRanges(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.paragraph#word-word-paragraph-gettextranges-member(1))|Gets the text ranges in the paragraph by using punctuation marks and/or other ending marks.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.paragraph#word-word-paragraph-inserttable-member(1))|Inserts a table with the specified number of rows and columns.|
||[isLastParagraph](/javascript/api/word/word.paragraph#word-word-paragraph-islastparagraph-member)|Indicates the paragraph is the last one inside its parent body.|
||[isListItem](/javascript/api/word/word.paragraph#word-word-paragraph-islistitem-member)|Checks whether the paragraph is a list item.|
||[list](/javascript/api/word/word.paragraph#word-word-paragraph-list-member)|Gets the List to which this paragraph belongs.|
||[listItem](/javascript/api/word/word.paragraph#word-word-paragraph-listitem-member)|Gets the ListItem for the paragraph.|
||[listItemOrNullObject](/javascript/api/word/word.paragraph#word-word-paragraph-listitemornullobject-member)|Gets the ListItem for the paragraph.|
||[listOrNullObject](/javascript/api/word/word.paragraph#word-word-paragraph-listornullobject-member)|Gets the List to which this paragraph belongs.|
||[parentBody](/javascript/api/word/word.paragraph#word-word-paragraph-parentbody-member)|Gets the parent body of the paragraph.|
||[parentContentControlOrNullObject](/javascript/api/word/word.paragraph#word-word-paragraph-parentcontentcontrolornullobject-member)|Gets the content control that contains the paragraph.|
||[parentTable](/javascript/api/word/word.paragraph#word-word-paragraph-parenttable-member)|Gets the table that contains the paragraph.|
||[parentTableCell](/javascript/api/word/word.paragraph#word-word-paragraph-parenttablecell-member)|Gets the table cell that contains the paragraph.|
||[parentTableCellOrNullObject](/javascript/api/word/word.paragraph#word-word-paragraph-parenttablecellornullobject-member)|Gets the table cell that contains the paragraph.|
||[parentTableOrNullObject](/javascript/api/word/word.paragraph#word-word-paragraph-parenttableornullobject-member)|Gets the table that contains the paragraph.|
||[split(delimiters: string[], trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.paragraph#word-word-paragraph-split-member(1))|Splits the paragraph into child ranges by using delimiters.|
||[startNewList()](/javascript/api/word/word.paragraph#word-word-paragraph-startnewlist-member(1))|Starts a new list with this paragraph.|
||[styleBuiltIn](/javascript/api/word/word.paragraph#word-word-paragraph-stylebuiltin-member)|Gets or sets the built-in style name for the paragraph.|
||[tableNestingLevel](/javascript/api/word/word.paragraph#word-word-paragraph-tablenestinglevel-member)|Gets the level of the paragraph's table.|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[getFirst()](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-getfirst-member(1))|Gets the first paragraph in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-getfirstornullobject-member(1))|Gets the first paragraph in this collection.|
||[getLast()](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-getlast-member(1))|Gets the last paragraph in this collection.|
||[getLastOrNullObject()](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-getlastornullobject-member(1))|Gets the last paragraph in this collection.|
|[Range](/javascript/api/word/word.range)|[compareLocationWith(range: Word.Range)](/javascript/api/word/word.range#word-word-range-comparelocationwith-member(1))|Compares this range's location with another range's location.|
||[expandTo(range: Word.Range)](/javascript/api/word/word.range#word-word-range-expandto-member(1))|Returns a new range that extends from this range in either direction to cover another range.|
||[expandToOrNullObject(range: Word.Range)](/javascript/api/word/word.range#word-word-range-expandtoornullobject-member(1))|Returns a new range that extends from this range in either direction to cover another range.|
||[getHyperlinkRanges()](/javascript/api/word/word.range#word-word-range-gethyperlinkranges-member(1))|Gets hyperlink child ranges within the range.|
||[getNextTextRange(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#word-word-range-getnexttextrange-member(1))|Gets the next text range by using punctuation marks and/or other ending marks.|
||[getNextTextRangeOrNullObject(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#word-word-range-getnexttextrangeornullobject-member(1))|Gets the next text range by using punctuation marks and/or other ending marks.|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.range#word-word-range-getrange-member(1))|Clones the range, or gets the starting or ending point of the range as a new range.|
||[getTextRanges(endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#word-word-range-gettextranges-member(1))|Gets the text child ranges in the range by using punctuation marks and/or other ending marks.|
||[hyperlink](/javascript/api/word/word.range#word-word-range-hyperlink-member)|Gets the first hyperlink in the range, or sets a hyperlink on the range.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.range#word-word-range-inserttable-member(1))|Inserts a table with the specified number of rows and columns.|
||[intersectWith(range: Word.Range)](/javascript/api/word/word.range#word-word-range-intersectwith-member(1))|Returns a new range as the intersection of this range with another range.|
||[intersectWithOrNullObject(range: Word.Range)](/javascript/api/word/word.range#word-word-range-intersectwithornullobject-member(1))|Returns a new range as the intersection of this range with another range.|
||[isEmpty](/javascript/api/word/word.range#word-word-range-isempty-member)|Checks whether the range length is zero.|
||[lists](/javascript/api/word/word.range#word-word-range-lists-member)|Gets the collection of list objects in the range.|
||[parentBody](/javascript/api/word/word.range#word-word-range-parentbody-member)|Gets the parent body of the range.|
||[parentContentControlOrNullObject](/javascript/api/word/word.range#word-word-range-parentcontentcontrolornullobject-member)|Gets the content control that contains the range.|
||[parentTable](/javascript/api/word/word.range#word-word-range-parenttable-member)|Gets the table that contains the range.|
||[parentTableCell](/javascript/api/word/word.range#word-word-range-parenttablecell-member)|Gets the table cell that contains the range.|
||[parentTableCellOrNullObject](/javascript/api/word/word.range#word-word-range-parenttablecellornullobject-member)|Gets the table cell that contains the range.|
||[parentTableOrNullObject](/javascript/api/word/word.range#word-word-range-parenttableornullobject-member)|Gets the table that contains the range.|
||[split(delimiters: string[], multiParagraphs?: boolean, trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.range#word-word-range-split-member(1))|Splits the range into child ranges by using delimiters.|
||[styleBuiltIn](/javascript/api/word/word.range#word-word-range-stylebuiltin-member)|Gets or sets the built-in style name for the range.|
||[tables](/javascript/api/word/word.range#word-word-range-tables-member)|Gets the collection of table objects in the range.|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[getFirst()](/javascript/api/word/word.rangecollection#word-word-rangecollection-getfirst-member(1))|Gets the first range in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.rangecollection#word-word-rangecollection-getfirstornullobject-member(1))|Gets the first range in this collection.|
|[RequestContext](/javascript/api/word/word.requestcontext)|[application](/javascript/api/word/word.requestcontext#word-word-requestcontext-application-member)|[Api set: WordApi 1.3] *|
|[Section](/javascript/api/word/word.section)|[getNext()](/javascript/api/word/word.section#word-word-section-getnext-member(1))|Gets the next section.|
||[getNextOrNullObject()](/javascript/api/word/word.section#word-word-section-getnextornullobject-member(1))|Gets the next section.|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[getFirst()](/javascript/api/word/word.sectioncollection#word-word-sectioncollection-getfirst-member(1))|Gets the first section in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.sectioncollection#word-word-sectioncollection-getfirstornullobject-member(1))|Gets the first section in this collection.|
|[Table](/javascript/api/word/word.table)|[addColumns(insertLocation: Word.InsertLocation, columnCount: number, values?: string[][])](/javascript/api/word/word.table#word-word-table-addcolumns-member(1))|Adds columns to the start or end of the table, using the first or last existing column as a template.|
||[addRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[][])](/javascript/api/word/word.table#word-word-table-addrows-member(1))|Adds rows to the start or end of the table, using the first or last existing row as a template.|
||[alignment](/javascript/api/word/word.table#word-word-table-alignment-member)|Gets or sets the alignment of the table against the page column.|
||[autoFitWindow()](/javascript/api/word/word.table#word-word-table-autofitwindow-member(1))|Autofits the table columns to the width of the window.|
||[clear()](/javascript/api/word/word.table#word-word-table-clear-member(1))|Clears the contents of the table.|
||[delete()](/javascript/api/word/word.table#word-word-table-delete-member(1))|Deletes the entire table.|
||[deleteColumns(columnIndex: number, columnCount?: number)](/javascript/api/word/word.table#word-word-table-deletecolumns-member(1))|Deletes specific columns.|
||[deleteRows(rowIndex: number, rowCount?: number)](/javascript/api/word/word.table#word-word-table-deleterows-member(1))|Deletes specific rows.|
||[distributeColumns()](/javascript/api/word/word.table#word-word-table-distributecolumns-member(1))|Distributes the column widths evenly.|
||[font](/javascript/api/word/word.table#word-word-table-font-member)|Gets the font.|
||[getBorder(borderLocation: Word.BorderLocation)](/javascript/api/word/word.table#word-word-table-getborder-member(1))|Gets the border style for the specified border.|
||[getCell(rowIndex: number, cellIndex: number)](/javascript/api/word/word.table#word-word-table-getcell-member(1))|Gets the table cell at a specified row and column.|
||[getCellOrNullObject(rowIndex: number, cellIndex: number)](/javascript/api/word/word.table#word-word-table-getcellornullobject-member(1))|Gets the table cell at a specified row and column.|
||[getCellPadding(cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.table#word-word-table-getcellpadding-member(1))|Gets cell padding in points.|
||[getNext()](/javascript/api/word/word.table#word-word-table-getnext-member(1))|Gets the next table.|
||[getNextOrNullObject()](/javascript/api/word/word.table#word-word-table-getnextornullobject-member(1))|Gets the next table.|
||[getParagraphAfter()](/javascript/api/word/word.table#word-word-table-getparagraphafter-member(1))|Gets the paragraph after the table.|
||[getParagraphAfterOrNullObject()](/javascript/api/word/word.table#word-word-table-getparagraphafterornullobject-member(1))|Gets the paragraph after the table.|
||[getParagraphBefore()](/javascript/api/word/word.table#word-word-table-getparagraphbefore-member(1))|Gets the paragraph before the table.|
||[getParagraphBeforeOrNullObject()](/javascript/api/word/word.table#word-word-table-getparagraphbeforeornullobject-member(1))|Gets the paragraph before the table.|
||[getRange(rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.table#word-word-table-getrange-member(1))|Gets the range that contains this table, or the range at the start or end of the table.|
||[headerRowCount](/javascript/api/word/word.table#word-word-table-headerrowcount-member)|Gets and sets the number of header rows.|
||[horizontalAlignment](/javascript/api/word/word.table#word-word-table-horizontalalignment-member)|Gets and sets the horizontal alignment of every cell in the table.|
||[ignorePunct](/javascript/api/word/word.table#word-word-table-ignorepunct-member)||
||[ignoreSpace](/javascript/api/word/word.table#word-word-table-ignorespace-member)||
||[insertContentControl()](/javascript/api/word/word.table#word-word-table-insertcontentcontrol-member(1))|Inserts a content control on the table.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.table#word-word-table-insertparagraph-member(1))|Inserts a paragraph at the specified location.|
||[insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.table#word-word-table-inserttable-member(1))|Inserts a table with the specified number of rows and columns.|
||[isUniform](/javascript/api/word/word.table#word-word-table-isuniform-member)|Indicates whether all of the table rows are uniform.|
||[matchCase](/javascript/api/word/word.table#word-word-table-matchcase-member)||
||[matchPrefix](/javascript/api/word/word.table#word-word-table-matchprefix-member)||
||[matchSuffix](/javascript/api/word/word.table#word-word-table-matchsuffix-member)||
||[matchWholeWord](/javascript/api/word/word.table#word-word-table-matchwholeword-member)||
||[matchWildcards](/javascript/api/word/word.table#word-word-table-matchwildcards-member)||
||[nestingLevel](/javascript/api/word/word.table#word-word-table-nestinglevel-member)|Gets the nesting level of the table.|
||[parentBody](/javascript/api/word/word.table#word-word-table-parentbody-member)|Gets the parent body of the table.|
||[parentContentControl](/javascript/api/word/word.table#word-word-table-parentcontentcontrol-member)|Gets the content control that contains the table.|
||[parentContentControlOrNullObject](/javascript/api/word/word.table#word-word-table-parentcontentcontrolornullobject-member)|Gets the content control that contains the table.|
||[parentTable](/javascript/api/word/word.table#word-word-table-parenttable-member)|Gets the table that contains this table.|
||[parentTableCell](/javascript/api/word/word.table#word-word-table-parenttablecell-member)|Gets the table cell that contains this table.|
||[parentTableCellOrNullObject](/javascript/api/word/word.table#word-word-table-parenttablecellornullobject-member)|Gets the table cell that contains this table.|
||[parentTableOrNullObject](/javascript/api/word/word.table#word-word-table-parenttableornullobject-member)|Gets the table that contains this table.|
||[rowCount](/javascript/api/word/word.table#word-word-table-rowcount-member)|Gets the number of rows in the table.|
||[rows](/javascript/api/word/word.table#word-word-table-rows-member)|Gets all of the table rows.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| {            ignorePunct?: boolean            ignoreSpace?: boolean            matchCase?: boolean            matchPrefix?: boolean            matchSuffix?: boolean            matchWholeWord?: boolean            matchWildcards?: boolean        })](/javascript/api/word/word.table#word-word-table-search-member(1))|Performs a search with the specified SearchOptions on the scope of the table object.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.table#word-word-table-select-member(1))|Selects the table, or the position at the start or end of the table, and navigates the Word UI to it.|
||[setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.table#word-word-table-setcellpadding-member(1))|Sets cell padding in points.|
||[shadingColor](/javascript/api/word/word.table#word-word-table-shadingcolor-member)|Gets and sets the shading color.|
||[style](/javascript/api/word/word.table#word-word-table-style-member)|Gets or sets the style name for the table.|
||[styleBandedColumns](/javascript/api/word/word.table#word-word-table-stylebandedcolumns-member)|Gets and sets whether the table has banded columns.|
||[styleBandedRows](/javascript/api/word/word.table#word-word-table-stylebandedrows-member)|Gets and sets whether the table has banded rows.|
||[styleBuiltIn](/javascript/api/word/word.table#word-word-table-stylebuiltin-member)|Gets or sets the built-in style name for the table.|
||[styleFirstColumn](/javascript/api/word/word.table#word-word-table-stylefirstcolumn-member)|Gets and sets whether the table has a first column with a special style.|
||[styleLastColumn](/javascript/api/word/word.table#word-word-table-stylelastcolumn-member)|Gets and sets whether the table has a last column with a special style.|
||[styleTotalRow](/javascript/api/word/word.table#word-word-table-styletotalrow-member)|Gets and sets whether the table has a total (last) row with a special style.|
||[tables](/javascript/api/word/word.table#word-word-table-tables-member)|Gets the child tables nested one level deeper.|
||[values](/javascript/api/word/word.table#word-word-table-values-member)|Gets and sets the text values in the table, as a 2D Javascript array.|
||[verticalAlignment](/javascript/api/word/word.table#word-word-table-verticalalignment-member)|Gets and sets the vertical alignment of every cell in the table.|
||[width](/javascript/api/word/word.table#word-word-table-width-member)|Gets and sets the width of the table in points.|
|[TableBorder](/javascript/api/word/word.tableborder)|[color](/javascript/api/word/word.tableborder#word-word-tableborder-color-member)|Gets or sets the table border color.|
||[type](/javascript/api/word/word.tableborder#word-word-tableborder-type-member)|Gets or sets the type of the table border.|
||[width](/javascript/api/word/word.tableborder#word-word-tableborder-width-member)|Gets or sets the width, in points, of the table border.|
|[TableCell](/javascript/api/word/word.tablecell)|[body](/javascript/api/word/word.tablecell#word-word-tablecell-body-member)|Gets the body object of the cell.|
||[cellIndex](/javascript/api/word/word.tablecell#word-word-tablecell-cellindex-member)|Gets the index of the cell in its row.|
||[columnWidth](/javascript/api/word/word.tablecell#word-word-tablecell-columnwidth-member)|Gets and sets the width of the cell's column in points.|
||[deleteColumn()](/javascript/api/word/word.tablecell#word-word-tablecell-deletecolumn-member(1))|Deletes the column containing this cell.|
||[deleteRow()](/javascript/api/word/word.tablecell#word-word-tablecell-deleterow-member(1))|Deletes the row containing this cell.|
||[getBorder(borderLocation: Word.BorderLocation)](/javascript/api/word/word.tablecell#word-word-tablecell-getborder-member(1))|Gets the border style for the specified border.|
||[getCellPadding(cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.tablecell#word-word-tablecell-getcellpadding-member(1))|Gets cell padding in points.|
||[getNext()](/javascript/api/word/word.tablecell#word-word-tablecell-getnext-member(1))|Gets the next cell.|
||[getNextOrNullObject()](/javascript/api/word/word.tablecell#word-word-tablecell-getnextornullobject-member(1))|Gets the next cell.|
||[horizontalAlignment](/javascript/api/word/word.tablecell#word-word-tablecell-horizontalalignment-member)|Gets and sets the horizontal alignment of the cell.|
||[insertColumns(insertLocation: Word.InsertLocation, columnCount: number, values?: string[][])](/javascript/api/word/word.tablecell#word-word-tablecell-insertcolumns-member(1))|Adds columns to the left or right of the cell, using the cell's column as a template.|
||[insertRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[][])](/javascript/api/word/word.tablecell#word-word-tablecell-insertrows-member(1))|Inserts rows above or below the cell, using the cell's row as a template.|
||[parentRow](/javascript/api/word/word.tablecell#word-word-tablecell-parentrow-member)|Gets the parent row of the cell.|
||[parentTable](/javascript/api/word/word.tablecell#word-word-tablecell-parenttable-member)|Gets the parent table of the cell.|
||[rowIndex](/javascript/api/word/word.tablecell#word-word-tablecell-rowindex-member)|Gets the index of the cell's row in the table.|
||[setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.tablecell#word-word-tablecell-setcellpadding-member(1))|Sets cell padding in points.|
||[shadingColor](/javascript/api/word/word.tablecell#word-word-tablecell-shadingcolor-member)|Gets or sets the shading color of the cell.|
||[value](/javascript/api/word/word.tablecell#word-word-tablecell-value-member)|Gets and sets the text of the cell.|
||[verticalAlignment](/javascript/api/word/word.tablecell#word-word-tablecell-verticalalignment-member)|Gets and sets the vertical alignment of the cell.|
||[width](/javascript/api/word/word.tablecell#word-word-tablecell-width-member)|Gets the width of the cell in points.|
|[TableCellCollection](/javascript/api/word/word.tablecellcollection)|[getFirst()](/javascript/api/word/word.tablecellcollection#word-word-tablecellcollection-getfirst-member(1))|Gets the first table cell in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.tablecellcollection#word-word-tablecellcollection-getfirstornullobject-member(1))|Gets the first table cell in this collection.|
||[items](/javascript/api/word/word.tablecellcollection#word-word-tablecellcollection-items-member)|Gets the loaded child items in this collection.|
|[TableCollection](/javascript/api/word/word.tablecollection)|[getFirst()](/javascript/api/word/word.tablecollection#word-word-tablecollection-getfirst-member(1))|Gets the first table in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.tablecollection#word-word-tablecollection-getfirstornullobject-member(1))|Gets the first table in this collection.|
||[items](/javascript/api/word/word.tablecollection#word-word-tablecollection-items-member)|Gets the loaded child items in this collection.|
|[TableRow](/javascript/api/word/word.tablerow)|[cellCount](/javascript/api/word/word.tablerow#word-word-tablerow-cellcount-member)|Gets the number of cells in the row.|
||[cells](/javascript/api/word/word.tablerow#word-word-tablerow-cells-member)|Gets cells.|
||[clear()](/javascript/api/word/word.tablerow#word-word-tablerow-clear-member(1))|Clears the contents of the row.|
||[delete()](/javascript/api/word/word.tablerow#word-word-tablerow-delete-member(1))|Deletes the entire row.|
||[font](/javascript/api/word/word.tablerow#word-word-tablerow-font-member)|Gets the font.|
||[getBorder(borderLocation: Word.BorderLocation)](/javascript/api/word/word.tablerow#word-word-tablerow-getborder-member(1))|Gets the border style of the cells in the row.|
||[getCellPadding(cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.tablerow#word-word-tablerow-getcellpadding-member(1))|Gets cell padding in points.|
||[getNext()](/javascript/api/word/word.tablerow#word-word-tablerow-getnext-member(1))|Gets the next row.|
||[getNextOrNullObject()](/javascript/api/word/word.tablerow#word-word-tablerow-getnextornullobject-member(1))|Gets the next row.|
||[horizontalAlignment](/javascript/api/word/word.tablerow#word-word-tablerow-horizontalalignment-member)|Gets and sets the horizontal alignment of every cell in the row.|
||[ignorePunct](/javascript/api/word/word.tablerow#word-word-tablerow-ignorepunct-member)||
||[ignoreSpace](/javascript/api/word/word.tablerow#word-word-tablerow-ignorespace-member)||
||[insertRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[][])](/javascript/api/word/word.tablerow#word-word-tablerow-insertrows-member(1))|Inserts rows using this row as a template.|
||[isHeader](/javascript/api/word/word.tablerow#word-word-tablerow-isheader-member)|Checks whether the row is a header row.|
||[matchCase](/javascript/api/word/word.tablerow#word-word-tablerow-matchcase-member)||
||[matchPrefix](/javascript/api/word/word.tablerow#word-word-tablerow-matchprefix-member)||
||[matchSuffix](/javascript/api/word/word.tablerow#word-word-tablerow-matchsuffix-member)||
||[matchWholeWord](/javascript/api/word/word.tablerow#word-word-tablerow-matchwholeword-member)||
||[matchWildcards](/javascript/api/word/word.tablerow#word-word-tablerow-matchwildcards-member)||
||[parentTable](/javascript/api/word/word.tablerow#word-word-tablerow-parenttable-member)|Gets parent table.|
||[preferredHeight](/javascript/api/word/word.tablerow#word-word-tablerow-preferredheight-member)|Gets and sets the preferred height of the row in points.|
||[rowIndex](/javascript/api/word/word.tablerow#word-word-tablerow-rowindex-member)|Gets the index of the row in its parent table.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| {            ignorePunct?: boolean            ignoreSpace?: boolean            matchCase?: boolean            matchPrefix?: boolean            matchSuffix?: boolean            matchWholeWord?: boolean            matchWildcards?: boolean        })](/javascript/api/word/word.tablerow#word-word-tablerow-search-member(1))|Performs a search with the specified SearchOptions on the scope of the row.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.tablerow#word-word-tablerow-select-member(1))|Selects the row and navigates the Word UI to it.|
||[setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.tablerow#word-word-tablerow-setcellpadding-member(1))|Sets cell padding in points.|
||[shadingColor](/javascript/api/word/word.tablerow#word-word-tablerow-shadingcolor-member)|Gets and sets the shading color.|
||[values](/javascript/api/word/word.tablerow#word-word-tablerow-values-member)|Gets and sets the text values in the row, as a 2D Javascript array.|
||[verticalAlignment](/javascript/api/word/word.tablerow#word-word-tablerow-verticalalignment-member)|Gets and sets the vertical alignment of the cells in the row.|
|[TableRowCollection](/javascript/api/word/word.tablerowcollection)|[getFirst()](/javascript/api/word/word.tablerowcollection#word-word-tablerowcollection-getfirst-member(1))|Gets the first row in this collection.|
||[getFirstOrNullObject()](/javascript/api/word/word.tablerowcollection#word-word-tablerowcollection-getfirstornullobject-member(1))|Gets the first row in this collection.|
||[items](/javascript/api/word/word.tablerowcollection#word-word-tablerowcollection-items-member)|Gets the loaded child items in this collection.|

## See also

- [Word JavaScript API Reference Documentation](/javascript/api/word)
- [Word JavaScript API requirement sets](word-api-requirement-sets.md)
