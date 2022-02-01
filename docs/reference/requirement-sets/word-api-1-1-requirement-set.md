---
title: Word JavaScript API requirement set 1.1
description: 'Details about the WordApi 1.1 requirement set'
ms.date: 11/01/2021
ms.prod: word
ms.localizationpriority: medium
---

# What's new in Word JavaScript API 1.1

WordApi 1.1 is the first requirement set of the Word JavaScript API. It's the only Word API requirement set supported by Word 2016.

## API list

The following table lists the APIs in Word JavaScript API requirement set 1.1. To view API reference documentation for all APIs supported by Word JavaScript API requirement set 1.1, see [Word APIs in requirement set 1.1](/javascript/api/word?view=word-js-1.1&preserve-view=true).

| Class | Fields | Description |
|:---|:---|:---|
|[Body](/javascript/api/word/word.body)|[clear()](/javascript/api/word/word.body#word-word-body-clear-member(1))|Clears the contents of the body object.|
||[contentControls](/javascript/api/word/word.body#word-word-body-contentControls-member)|Gets the collection of rich text content control objects in the body.|
||[font](/javascript/api/word/word.body#word-word-body-font-member)|Gets the text format of the body.|
||[getHtml()](/javascript/api/word/word.body#word-word-body-getHtml-member(1))|Gets an HTML representation of the body object.|
||[getOoxml()](/javascript/api/word/word.body#word-word-body-getOoxml-member(1))|Gets the OOXML (Office Open XML) representation of the body object.|
||[ignorePunct](/javascript/api/word/word.body#word-word-body-ignorePunct-member)||
||[ignoreSpace](/javascript/api/word/word.body#word-word-body-ignoreSpace-member)||
||[inlinePictures](/javascript/api/word/word.body#word-word-body-inlinePictures-member)|Gets the collection of InlinePicture objects in the body.|
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#word-word-body-insertBreak-member(1))|Inserts a break at the specified location in the main document.|
||[insertContentControl()](/javascript/api/word/word.body#word-word-body-insertContentControl-member(1))|Wraps the body object with a Rich Text content control.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#insertFileFromBase64_base64File__insertLocation_)|Inserts a document into the body at the specified location.|
||[insertHtml(html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#word-word-body-insertHtml-member(1))|Inserts HTML at the specified location.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#word-word-body-insertOoxml-member(1))|Inserts OOXML at the specified location.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#word-word-body-insertParagraph-member(1))|Inserts a paragraph at the specified location.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#word-word-body-insertText-member(1))|Inserts text into the body at the specified location.|
||[matchCase](/javascript/api/word/word.body#word-word-body-matchCase-member)||
||[matchPrefix](/javascript/api/word/word.body#word-word-body-matchPrefix-member)||
||[matchSuffix](/javascript/api/word/word.body#word-word-body-matchSuffix-member)||
||[matchWholeWord](/javascript/api/word/word.body#word-word-body-matchWholeWord-member)||
||[matchWildcards](/javascript/api/word/word.body#word-word-body-matchWildcards-member)||
||[paragraphs](/javascript/api/word/word.body#word-word-body-paragraphs-member)|Gets the collection of paragraph objects in the body.|
||[parentContentControl](/javascript/api/word/word.body#word-word-body-parentContentControl-member)|Gets the content control that contains the body.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| {            ignorePunct?: boolean            ignoreSpace?: boolean            matchCase?: boolean            matchPrefix?: boolean            matchSuffix?: boolean            matchWholeWord?: boolean            matchWildcards?: boolean        })](/javascript/api/word/word.body#word-word-body-search-member(1))|Performs a search with the specified SearchOptions on the scope of the body object.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.body#word-word-body-select-member(1))|Selects the body and navigates the Word UI to it.|
||[style](/javascript/api/word/word.body#word-word-body-style-member)|Gets or sets the style name for the body.|
||[text](/javascript/api/word/word.body#word-word-body-text-member)|Gets the text of the body.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[appearance](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-appearance-member)|Gets or sets the appearance of the content control.|
||[cannotDelete](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-cannotDelete-member)|Gets or sets a value that indicates whether the user can delete the content control.|
||[cannotEdit](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-cannotEdit-member)|Gets or sets a value that indicates whether the user can edit the contents of the content control.|
||[clear()](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-clear-member(1))|Clears the contents of the content control.|
||[color](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-color-member)|Gets or sets the color of the content control.|
||[contentControls](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-contentControls-member)|Gets the collection of content control objects in the content control.|
||[delete(keepContent: boolean)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-delete-member(1))|Deletes the content control and its content.|
||[font](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-font-member)|Gets the text format of the content control.|
||[getHtml()](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-getHtml-member(1))|Gets an HTML representation of the content control object.|
||[getOoxml()](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-getOoxml-member(1))|Gets the Office Open XML (OOXML) representation of the content control object.|
||[id](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-id-member)|Gets an integer that represents the content control identifier.|
||[ignorePunct](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-ignorePunct-member)||
||[ignoreSpace](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-ignoreSpace-member)||
||[inlinePictures](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-inlinePictures-member)|Gets the collection of inlinePicture objects in the content control.|
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-insertBreak-member(1))|Inserts a break at the specified location in the main document.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#insertFileFromBase64_base64File__insertLocation_)|Inserts a document into the content control at the specified location.|
||[insertHtml(html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-insertHtml-member(1))|Inserts HTML into the content control at the specified location.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-insertOoxml-member(1))|Inserts OOXML into the content control at the specified location.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-insertParagraph-member(1))|Inserts a paragraph at the specified location.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-insertText-member(1))|Inserts text into the content control at the specified location.|
||[matchCase](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-matchCase-member)||
||[matchPrefix](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-matchPrefix-member)||
||[matchSuffix](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-matchSuffix-member)||
||[matchWholeWord](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-matchWholeWord-member)||
||[matchWildcards](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-matchWildcards-member)||
||[paragraphs](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-paragraphs-member)|Gets the collection of paragraph objects in the content control.|
||[parentContentControl](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parentContentControl-member)|Gets the content control that contains the content control.|
||[placeholderText](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-placeholderText-member)|Gets or sets the placeholder text of the content control.|
||[removeWhenEdited](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-removeWhenEdited-member)|Gets or sets a value that indicates whether the content control is removed after it is edited.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| {            ignorePunct?: boolean            ignoreSpace?: boolean            matchCase?: boolean            matchPrefix?: boolean            matchSuffix?: boolean            matchWholeWord?: boolean            matchWildcards?: boolean        })](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-search-member(1))|Performs a search with the specified SearchOptions on the scope of the content control object.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-select-member(1))|Selects the content control.|
||[style](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-style-member)|Gets or sets the style name for the content control.|
||[tag](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-tag-member)|Gets or sets a tag to identify a content control.|
||[text](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-text-member)|Gets the text of the content control.|
||[title](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-title-member)|Gets or sets the title for a content control.|
||[type](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-type-member)|Gets the content control type.|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getById(id: number)](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getById-member(1))|Gets a content control by its identifier.|
||[getByTag(tag: string)](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getByTag-member(1))|Gets the content controls that have the specified tag.|
||[getByTitle(title: string)](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getByTitle-member(1))|Gets the content controls that have the specified title.|
||[getItem(index: number)](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getItem-member(1))|Gets a content control by its index in the collection.|
||[items](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-items-member)|Gets the loaded child items in this collection.|
|[Document](/javascript/api/word/word.document)|[body](/javascript/api/word/word.document#word-word-document-body-member)|Gets the body object of the main document.|
||[contentControls](/javascript/api/word/word.document#word-word-document-contentControls-member)|Gets the collection of content control objects in the document.|
||[getSelection()](/javascript/api/word/word.document#word-word-document-getSelection-member(1))|Gets the current selection of the document.|
||[save()](/javascript/api/word/word.document#word-word-document-save-member(1))|Saves the document.|
||[saved](/javascript/api/word/word.document#word-word-document-saved-member)|Indicates whether the changes in the document have been saved.|
||[sections](/javascript/api/word/word.document#word-word-document-sections-member)|Gets the collection of section objects in the document.|
|[Font](/javascript/api/word/word.font)|[bold](/javascript/api/word/word.font#word-word-font-bold-member)|Gets or sets a value that indicates whether the font is bold.|
||[color](/javascript/api/word/word.font#word-word-font-color-member)|Gets or sets the color for the specified font.|
||[doubleStrikeThrough](/javascript/api/word/word.font#word-word-font-doubleStrikeThrough-member)|Gets or sets a value that indicates whether the font has a double strikethrough.|
||[highlightColor](/javascript/api/word/word.font#word-word-font-highlightColor-member)|Gets or sets the highlight color.|
||[italic](/javascript/api/word/word.font#word-word-font-italic-member)|Gets or sets a value that indicates whether the font is italicized.|
||[name](/javascript/api/word/word.font#word-word-font-name-member)|Gets or sets a value that represents the name of the font.|
||[size](/javascript/api/word/word.font#word-word-font-size-member)|Gets or sets a value that represents the font size in points.|
||[strikeThrough](/javascript/api/word/word.font#word-word-font-strikeThrough-member)|Gets or sets a value that indicates whether the font has a strikethrough.|
||[subscript](/javascript/api/word/word.font#word-word-font-subscript-member)|Gets or sets a value that indicates whether the font is a subscript.|
||[superscript](/javascript/api/word/word.font#word-word-font-superscript-member)|Gets or sets a value that indicates whether the font is a superscript.|
||[underline](/javascript/api/word/word.font#word-word-font-underline-member)|Gets or sets a value that indicates the font's underline type.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[altTextDescription](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-altTextDescription-member)|Gets or sets a string that represents the alternative text associated with the inline image.|
||[altTextTitle](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-altTextTitle-member)|Gets or sets a string that contains the title for the inline image.|
||[getBase64ImageSrc()](/javascript/api/word/word.inlinepicture#getBase64ImageSrc__)|Gets the base64 encoded string representation of the inline image.|
||[height](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-height-member)|Gets or sets a number that describes the height of the inline image.|
||[hyperlink](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-hyperlink-member)|Gets or sets a hyperlink on the image.|
||[insertContentControl()](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-insertContentControl-member(1))|Wraps the inline picture with a rich text content control.|
||[lockAspectRatio](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-lockAspectRatio-member)|Gets or sets a value that indicates whether the inline image retains its original proportions when you resize it.|
||[parentContentControl](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-parentContentControl-member)|Gets the content control that contains the inline image.|
||[width](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-width-member)|Gets or sets a number that describes the width of the inline image.|
|[InlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|[items](/javascript/api/word/word.inlinepicturecollection#word-word-inlinepicturecollection-items-member)|Gets the loaded child items in this collection.|
|[Paragraph](/javascript/api/word/word.paragraph)|[alignment](/javascript/api/word/word.paragraph#word-word-paragraph-alignment-member)|Gets or sets the alignment for a paragraph.|
||[clear()](/javascript/api/word/word.paragraph#word-word-paragraph-clear-member(1))|Clears the contents of the paragraph object.|
||[contentControls](/javascript/api/word/word.paragraph#word-word-paragraph-contentControls-member)|Gets the collection of content control objects in the paragraph.|
||[delete()](/javascript/api/word/word.paragraph#word-word-paragraph-delete-member(1))|Deletes the paragraph and its content from the document.|
||[firstLineIndent](/javascript/api/word/word.paragraph#word-word-paragraph-firstLineIndent-member)|Gets or sets the value, in points, for a first line or hanging indent.|
||[font](/javascript/api/word/word.paragraph#word-word-paragraph-font-member)|Gets the text format of the paragraph.|
||[getHtml()](/javascript/api/word/word.paragraph#word-word-paragraph-getHtml-member(1))|Gets an HTML representation of the paragraph object.|
||[getOoxml()](/javascript/api/word/word.paragraph#word-word-paragraph-getOoxml-member(1))|Gets the Office Open XML (OOXML) representation of the paragraph object.|
||[ignorePunct](/javascript/api/word/word.paragraph#word-word-paragraph-ignorePunct-member)||
||[ignoreSpace](/javascript/api/word/word.paragraph#word-word-paragraph-ignoreSpace-member)||
||[inlinePictures](/javascript/api/word/word.paragraph#word-word-paragraph-inlinePictures-member)|Gets the collection of InlinePicture objects in the paragraph.|
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#word-word-paragraph-insertBreak-member(1))|Inserts a break at the specified location in the main document.|
||[insertContentControl()](/javascript/api/word/word.paragraph#word-word-paragraph-insertContentControl-member(1))|Wraps the paragraph object with a rich text content control.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#insertFileFromBase64_base64File__insertLocation_)|Inserts a document into the paragraph at the specified location.|
||[insertHtml(html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#word-word-paragraph-insertHtml-member(1))|Inserts HTML into the paragraph at the specified location.|
||[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#insertInlinePictureFromBase64_base64EncodedImage__insertLocation_)|Inserts a picture into the paragraph at the specified location.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#word-word-paragraph-insertOoxml-member(1))|Inserts OOXML into the paragraph at the specified location.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#word-word-paragraph-insertParagraph-member(1))|Inserts a paragraph at the specified location.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#word-word-paragraph-insertText-member(1))|Inserts text into the paragraph at the specified location.|
||[leftIndent](/javascript/api/word/word.paragraph#word-word-paragraph-leftIndent-member)|Gets or sets the left indent value, in points, for the paragraph.|
||[lineSpacing](/javascript/api/word/word.paragraph#word-word-paragraph-lineSpacing-member)|Gets or sets the line spacing, in points, for the specified paragraph.|
||[lineUnitAfter](/javascript/api/word/word.paragraph#word-word-paragraph-lineUnitAfter-member)|Gets or sets the amount of spacing, in grid lines, after the paragraph.|
||[lineUnitBefore](/javascript/api/word/word.paragraph#word-word-paragraph-lineUnitBefore-member)|Gets or sets the amount of spacing, in grid lines, before the paragraph.|
||[matchCase](/javascript/api/word/word.paragraph#word-word-paragraph-matchCase-member)||
||[matchPrefix](/javascript/api/word/word.paragraph#word-word-paragraph-matchPrefix-member)||
||[matchSuffix](/javascript/api/word/word.paragraph#word-word-paragraph-matchSuffix-member)||
||[matchWholeWord](/javascript/api/word/word.paragraph#word-word-paragraph-matchWholeWord-member)||
||[matchWildcards](/javascript/api/word/word.paragraph#word-word-paragraph-matchWildcards-member)||
||[outlineLevel](/javascript/api/word/word.paragraph#word-word-paragraph-outlineLevel-member)|Gets or sets the outline level for the paragraph.|
||[parentContentControl](/javascript/api/word/word.paragraph#word-word-paragraph-parentContentControl-member)|Gets the content control that contains the paragraph.|
||[rightIndent](/javascript/api/word/word.paragraph#word-word-paragraph-rightIndent-member)|Gets or sets the right indent value, in points, for the paragraph.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| {            ignorePunct?: boolean            ignoreSpace?: boolean            matchCase?: boolean            matchPrefix?: boolean            matchSuffix?: boolean            matchWholeWord?: boolean            matchWildcards?: boolean        })](/javascript/api/word/word.paragraph#word-word-paragraph-search-member(1))|Performs a search with the specified SearchOptions on the scope of the paragraph object.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.paragraph#word-word-paragraph-select-member(1))|Selects and navigates the Word UI to the paragraph.|
||[spaceAfter](/javascript/api/word/word.paragraph#word-word-paragraph-spaceAfter-member)|Gets or sets the spacing, in points, after the paragraph.|
||[spaceBefore](/javascript/api/word/word.paragraph#word-word-paragraph-spaceBefore-member)|Gets or sets the spacing, in points, before the paragraph.|
||[style](/javascript/api/word/word.paragraph#word-word-paragraph-style-member)|Gets or sets the style name for the paragraph.|
||[text](/javascript/api/word/word.paragraph#word-word-paragraph-text-member)|Gets the text of the paragraph.|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[items](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-items-member)|Gets the loaded child items in this collection.|
|[Range](/javascript/api/word/word.range)|[clear()](/javascript/api/word/word.range#word-word-range-clear-member(1))|Clears the contents of the range object.|
||[contentControls](/javascript/api/word/word.range#word-word-range-contentControls-member)|Gets the collection of content control objects in the range.|
||[delete()](/javascript/api/word/word.range#word-word-range-delete-member(1))|Deletes the range and its content from the document.|
||[font](/javascript/api/word/word.range#word-word-range-font-member)|Gets the text format of the range.|
||[getHtml()](/javascript/api/word/word.range#word-word-range-getHtml-member(1))|Gets an HTML representation of the range object.|
||[getOoxml()](/javascript/api/word/word.range#word-word-range-getOoxml-member(1))|Gets the OOXML representation of the range object.|
||[ignorePunct](/javascript/api/word/word.range#word-word-range-ignorePunct-member)||
||[ignoreSpace](/javascript/api/word/word.range#word-word-range-ignoreSpace-member)||
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#word-word-range-insertBreak-member(1))|Inserts a break at the specified location in the main document.|
||[insertContentControl()](/javascript/api/word/word.range#word-word-range-insertContentControl-member(1))|Wraps the range object with a rich text content control.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#insertFileFromBase64_base64File__insertLocation_)|Inserts a document at the specified location.|
||[insertHtml(html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#word-word-range-insertHtml-member(1))|Inserts HTML at the specified location.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#word-word-range-insertOoxml-member(1))|Inserts OOXML at the specified location.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#word-word-range-insertParagraph-member(1))|Inserts a paragraph at the specified location.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#word-word-range-insertText-member(1))|Inserts text at the specified location.|
||[matchCase](/javascript/api/word/word.range#word-word-range-matchCase-member)||
||[matchPrefix](/javascript/api/word/word.range#word-word-range-matchPrefix-member)||
||[matchSuffix](/javascript/api/word/word.range#word-word-range-matchSuffix-member)||
||[matchWholeWord](/javascript/api/word/word.range#word-word-range-matchWholeWord-member)||
||[matchWildcards](/javascript/api/word/word.range#word-word-range-matchWildcards-member)||
||[paragraphs](/javascript/api/word/word.range#word-word-range-paragraphs-member)|Gets the collection of paragraph objects in the range.|
||[parentContentControl](/javascript/api/word/word.range#word-word-range-parentContentControl-member)|Gets the content control that contains the range.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| {            ignorePunct?: boolean            ignoreSpace?: boolean            matchCase?: boolean            matchPrefix?: boolean            matchSuffix?: boolean            matchWholeWord?: boolean            matchWildcards?: boolean        })](/javascript/api/word/word.range#word-word-range-search-member(1))|Performs a search with the specified SearchOptions on the scope of the range object.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.range#word-word-range-select-member(1))|Selects and navigates the Word UI to the range.|
||[style](/javascript/api/word/word.range#word-word-range-style-member)|Gets or sets the style name for the range.|
||[text](/javascript/api/word/word.range#word-word-range-text-member)|Gets the text of the range.|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[items](/javascript/api/word/word.rangecollection#word-word-rangecollection-items-member)|Gets the loaded child items in this collection.|
|[SearchOptions](/javascript/api/word/word.searchoptions)|[ignorePunct](/javascript/api/word/word.searchoptions#word-word-searchoptions-ignorePunct-member)|Gets or sets a value that indicates whether to ignore all punctuation characters between words.|
||[ignoreSpace](/javascript/api/word/word.searchoptions#word-word-searchoptions-ignoreSpace-member)|Gets or sets a value that indicates whether to ignore all whitespace between words.|
||[matchCase](/javascript/api/word/word.searchoptions#word-word-searchoptions-matchCase-member)|Gets or sets a value that indicates whether to perform a case sensitive search.|
||[matchPrefix](/javascript/api/word/word.searchoptions#word-word-searchoptions-matchPrefix-member)|Gets or sets a value that indicates whether to match words that begin with the search string.|
||[matchSuffix](/javascript/api/word/word.searchoptions#word-word-searchoptions-matchSuffix-member)|Gets or sets a value that indicates whether to match words that end with the search string.|
||[matchWholeWord](/javascript/api/word/word.searchoptions#word-word-searchoptions-matchWholeWord-member)|Gets or sets a value that indicates whether to find operation only entire words, not text that is part of a larger word.|
||[matchWildcards](/javascript/api/word/word.searchoptions#word-word-searchoptions-matchWildcards-member)|Gets or sets a value that indicates whether the search will be performed using special search operators.|
|[Section](/javascript/api/word/word.section)|[body](/javascript/api/word/word.section#word-word-section-body-member)|Gets the body object of the section.|
||[getFooter(type: Word.HeaderFooterType)](/javascript/api/word/word.section#word-word-section-getFooter-member(1))|Gets one of the section's footers.|
||[getHeader(type: Word.HeaderFooterType)](/javascript/api/word/word.section#word-word-section-getHeader-member(1))|Gets one of the section's headers.|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[items](/javascript/api/word/word.sectioncollection#word-word-sectioncollection-items-member)|Gets the loaded child items in this collection.|

## See also

- [Word JavaScript API Reference Documentation](/javascript/api/word)
- [Word JavaScript API requirement sets](word-api-requirement-sets.md)
