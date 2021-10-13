---
title: Word JavaScript API requirement set 1.1
description: 'Details about the WordApi 1.1 requirement set'
ms.date: 11/09/2020
ms.prod: word
ms.localizationpriority: medium
---

# What's new in Word JavaScript API 1.1

WordApi 1.1 is the first requirement set of the Word JavaScript API. It's the only Word API requirement set supported by Word 2016.

## API list

The following table lists the APIs in Word JavaScript API requirement set 1.1. To view API reference documentation for all APIs supported by Word JavaScript API requirement set 1.1, see [Word APIs in requirement set 1.1](/javascript/api/word?view=word-js-1.1&preserve-view=true).

| Class | Fields | Description |
|:---|:---|:---|
|[Body](/javascript/api/word/word.body)|[clear()](/javascript/api/word/word.body#clear__)|Clears the contents of the body object.|
||[contentControls](/javascript/api/word/word.body#contentControls)|Gets the collection of rich text content control objects in the body.|
||[font](/javascript/api/word/word.body#font)|Gets the text format of the body.|
||[getHtml()](/javascript/api/word/word.body#getHtml__)|Gets an HTML representation of the body object.|
||[getOoxml()](/javascript/api/word/word.body#getOoxml__)|Gets the OOXML (Office Open XML) representation of the body object.|
||[ignorePunct](/javascript/api/word/word.body#ignorePunct)||
||[ignoreSpace](/javascript/api/word/word.body#ignoreSpace)||
||[inlinePictures](/javascript/api/word/word.body#inlinePictures)|Gets the collection of InlinePicture objects in the body.|
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#insertBreak_breakType__insertLocation_)|Inserts a break at the specified location in the main document.|
||[insertContentControl()](/javascript/api/word/word.body#insertContentControl__)|Wraps the body object with a Rich Text content control.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#insertFileFromBase64_base64File__insertLocation_)|Inserts a document into the body at the specified location.|
||[insertHtml(html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#insertHtml_html__insertLocation_)|Inserts HTML at the specified location.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#insertOoxml_ooxml__insertLocation_)|Inserts OOXML at the specified location.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#insertParagraph_paragraphText__insertLocation_)|Inserts a paragraph at the specified location.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#insertText_text__insertLocation_)|Inserts text into the body at the specified location.|
||[matchCase](/javascript/api/word/word.body#matchCase)||
||[matchPrefix](/javascript/api/word/word.body#matchPrefix)||
||[matchSuffix](/javascript/api/word/word.body#matchSuffix)||
||[matchWholeWord](/javascript/api/word/word.body#matchWholeWord)||
||[matchWildcards](/javascript/api/word/word.body#matchWildcards)||
||[paragraphs](/javascript/api/word/word.body#paragraphs)|Gets the collection of paragraph objects in the body.|
||[parentContentControl](/javascript/api/word/word.body#parentContentControl)|Gets the content control that contains the body.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| {            ignorePunct?: boolean            ignoreSpace?: boolean            matchCase?: boolean            matchPrefix?: boolean            matchSuffix?: boolean            matchWholeWord?: boolean            matchWildcards?: boolean        })](/javascript/api/word/word.body#search_searchText__searchOptions__ignorePunct__ignoreSpace__matchCase__matchPrefix__matchSuffix__matchWholeWord__matchWildcards_)|Performs a search with the specified SearchOptions on the scope of the body object.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.body#select_selectionMode_)|Selects the body and navigates the Word UI to it.|
||[style](/javascript/api/word/word.body#style)|Gets or sets the style name for the body.|
||[text](/javascript/api/word/word.body#text)|Gets the text of the body.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[appearance](/javascript/api/word/word.contentcontrol#appearance)|Gets or sets the appearance of the content control.|
||[cannotDelete](/javascript/api/word/word.contentcontrol#cannotDelete)|Gets or sets a value that indicates whether the user can delete the content control.|
||[cannotEdit](/javascript/api/word/word.contentcontrol#cannotEdit)|Gets or sets a value that indicates whether the user can edit the contents of the content control.|
||[clear()](/javascript/api/word/word.contentcontrol#clear__)|Clears the contents of the content control.|
||[color](/javascript/api/word/word.contentcontrol#color)|Gets or sets the color of the content control.|
||[contentControls](/javascript/api/word/word.contentcontrol#contentControls)|Gets the collection of content control objects in the content control.|
||[delete(keepContent: boolean)](/javascript/api/word/word.contentcontrol#delete_keepContent_)|Deletes the content control and its content.|
||[font](/javascript/api/word/word.contentcontrol#font)|Gets the text format of the content control.|
||[getHtml()](/javascript/api/word/word.contentcontrol#getHtml__)|Gets an HTML representation of the content control object.|
||[getOoxml()](/javascript/api/word/word.contentcontrol#getOoxml__)|Gets the Office Open XML (OOXML) representation of the content control object.|
||[id](/javascript/api/word/word.contentcontrol#id)|Gets an integer that represents the content control identifier.|
||[ignorePunct](/javascript/api/word/word.contentcontrol#ignorePunct)||
||[ignoreSpace](/javascript/api/word/word.contentcontrol#ignoreSpace)||
||[inlinePictures](/javascript/api/word/word.contentcontrol#inlinePictures)|Gets the collection of inlinePicture objects in the content control.|
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#insertBreak_breakType__insertLocation_)|Inserts a break at the specified location in the main document.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#insertFileFromBase64_base64File__insertLocation_)|Inserts a document into the content control at the specified location.|
||[insertHtml(html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#insertHtml_html__insertLocation_)|Inserts HTML into the content control at the specified location.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#insertOoxml_ooxml__insertLocation_)|Inserts OOXML into the content control at the specified location.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#insertParagraph_paragraphText__insertLocation_)|Inserts a paragraph at the specified location.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#insertText_text__insertLocation_)|Inserts text into the content control at the specified location.|
||[matchCase](/javascript/api/word/word.contentcontrol#matchCase)||
||[matchPrefix](/javascript/api/word/word.contentcontrol#matchPrefix)||
||[matchSuffix](/javascript/api/word/word.contentcontrol#matchSuffix)||
||[matchWholeWord](/javascript/api/word/word.contentcontrol#matchWholeWord)||
||[matchWildcards](/javascript/api/word/word.contentcontrol#matchWildcards)||
||[paragraphs](/javascript/api/word/word.contentcontrol#paragraphs)|Get the collection of paragraph objects in the content control.|
||[parentContentControl](/javascript/api/word/word.contentcontrol#parentContentControl)|Gets the content control that contains the content control.|
||[placeholderText](/javascript/api/word/word.contentcontrol#placeholderText)|Gets or sets the placeholder text of the content control.|
||[removeWhenEdited](/javascript/api/word/word.contentcontrol#removeWhenEdited)|Gets or sets a value that indicates whether the content control is removed after it is edited.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| {            ignorePunct?: boolean            ignoreSpace?: boolean            matchCase?: boolean            matchPrefix?: boolean            matchSuffix?: boolean            matchWholeWord?: boolean            matchWildcards?: boolean        })](/javascript/api/word/word.contentcontrol#search_searchText__searchOptions__ignorePunct__ignoreSpace__matchCase__matchPrefix__matchSuffix__matchWholeWord__matchWildcards_)|Performs a search with the specified SearchOptions on the scope of the content control object.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.contentcontrol#select_selectionMode_)|Selects the content control.|
||[style](/javascript/api/word/word.contentcontrol#style)|Gets or sets the style name for the content control.|
||[tag](/javascript/api/word/word.contentcontrol#tag)|Gets or sets a tag to identify a content control.|
||[text](/javascript/api/word/word.contentcontrol#text)|Gets the text of the content control.|
||[title](/javascript/api/word/word.contentcontrol#title)|Gets or sets the title for a content control.|
||[type](/javascript/api/word/word.contentcontrol#type)|Gets the content control type.|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getById(id: number)](/javascript/api/word/word.contentcontrolcollection#getById_id_)|Gets a content control by its identifier.|
||[getByTag(tag: string)](/javascript/api/word/word.contentcontrolcollection#getByTag_tag_)|Gets the content controls that have the specified tag.|
||[getByTitle(title: string)](/javascript/api/word/word.contentcontrolcollection#getByTitle_title_)|Gets the content controls that have the specified title.|
||[getItem(index: number)](/javascript/api/word/word.contentcontrolcollection#getItem_index_)|Gets a content control by its index in the collection.|
||[items](/javascript/api/word/word.contentcontrolcollection#items)|Gets the loaded child items in this collection.|
|[Document](/javascript/api/word/word.document)|[body](/javascript/api/word/word.document#body)|Gets the body object of the main document.|
||[contentControls](/javascript/api/word/word.document#contentControls)|Gets the collection of content control objects in the document.|
||[getSelection()](/javascript/api/word/word.document#getSelection__)|Gets the current selection of the document.|
||[save()](/javascript/api/word/word.document#save__)|Saves the document.|
||[saved](/javascript/api/word/word.document#saved)|Indicates whether the changes in the document have been saved.|
||[sections](/javascript/api/word/word.document#sections)|Gets the collection of section objects in the document.|
|[Font](/javascript/api/word/word.font)|[bold](/javascript/api/word/word.font#bold)|Gets or sets a value that indicates whether the font is bold.|
||[color](/javascript/api/word/word.font#color)|Gets or sets the color for the specified font.|
||[doubleStrikeThrough](/javascript/api/word/word.font#doubleStrikeThrough)|Gets or sets a value that indicates whether the font has a double strikethrough.|
||[highlightColor](/javascript/api/word/word.font#highlightColor)|Gets or sets the highlight color.|
||[italic](/javascript/api/word/word.font#italic)|Gets or sets a value that indicates whether the font is italicized.|
||[name](/javascript/api/word/word.font#name)|Gets or sets a value that represents the name of the font.|
||[size](/javascript/api/word/word.font#size)|Gets or sets a value that represents the font size in points.|
||[strikeThrough](/javascript/api/word/word.font#strikeThrough)|Gets or sets a value that indicates whether the font has a strikethrough.|
||[subscript](/javascript/api/word/word.font#subscript)|Gets or sets a value that indicates whether the font is a subscript.|
||[superscript](/javascript/api/word/word.font#superscript)|Gets or sets a value that indicates whether the font is a superscript.|
||[underline](/javascript/api/word/word.font#underline)|Gets or sets a value that indicates the font's underline type.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[altTextDescription](/javascript/api/word/word.inlinepicture#altTextDescription)|Gets or sets a string that represents the alternative text associated with the inline image.|
||[altTextTitle](/javascript/api/word/word.inlinepicture#altTextTitle)|Gets or sets a string that contains the title for the inline image.|
||[getBase64ImageSrc()](/javascript/api/word/word.inlinepicture#getBase64ImageSrc__)|Gets the base64 encoded string representation of the inline image.|
||[height](/javascript/api/word/word.inlinepicture#height)|Gets or sets a number that describes the height of the inline image.|
||[hyperlink](/javascript/api/word/word.inlinepicture#hyperlink)|Gets or sets a hyperlink on the image.|
||[insertContentControl()](/javascript/api/word/word.inlinepicture#insertContentControl__)|Wraps the inline picture with a rich text content control.|
||[lockAspectRatio](/javascript/api/word/word.inlinepicture#lockAspectRatio)|Gets or sets a value that indicates whether the inline image retains its original proportions when you resize it.|
||[parentContentControl](/javascript/api/word/word.inlinepicture#parentContentControl)|Gets the content control that contains the inline image.|
||[width](/javascript/api/word/word.inlinepicture#width)|Gets or sets a number that describes the width of the inline image.|
|[InlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|[items](/javascript/api/word/word.inlinepicturecollection#items)|Gets the loaded child items in this collection.|
|[Paragraph](/javascript/api/word/word.paragraph)|[alignment](/javascript/api/word/word.paragraph#alignment)|Gets or sets the alignment for a paragraph.|
||[clear()](/javascript/api/word/word.paragraph#clear__)|Clears the contents of the paragraph object.|
||[contentControls](/javascript/api/word/word.paragraph#contentControls)|Gets the collection of content control objects in the paragraph.|
||[delete()](/javascript/api/word/word.paragraph#delete__)|Deletes the paragraph and its content from the document.|
||[firstLineIndent](/javascript/api/word/word.paragraph#firstLineIndent)|Gets or sets the value, in points, for a first line or hanging indent.|
||[font](/javascript/api/word/word.paragraph#font)|Gets the text format of the paragraph.|
||[getHtml()](/javascript/api/word/word.paragraph#getHtml__)|Gets an HTML representation of the paragraph object.|
||[getOoxml()](/javascript/api/word/word.paragraph#getOoxml__)|Gets the Office Open XML (OOXML) representation of the paragraph object.|
||[ignorePunct](/javascript/api/word/word.paragraph#ignorePunct)||
||[ignoreSpace](/javascript/api/word/word.paragraph#ignoreSpace)||
||[inlinePictures](/javascript/api/word/word.paragraph#inlinePictures)|Gets the collection of InlinePicture objects in the paragraph.|
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#insertBreak_breakType__insertLocation_)|Inserts a break at the specified location in the main document.|
||[insertContentControl()](/javascript/api/word/word.paragraph#insertContentControl__)|Wraps the paragraph object with a rich text content control.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#insertFileFromBase64_base64File__insertLocation_)|Inserts a document into the paragraph at the specified location.|
||[insertHtml(html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#insertHtml_html__insertLocation_)|Inserts HTML into the paragraph at the specified location.|
||[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#insertInlinePictureFromBase64_base64EncodedImage__insertLocation_)|Inserts a picture into the paragraph at the specified location.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#insertOoxml_ooxml__insertLocation_)|Inserts OOXML into the paragraph at the specified location.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#insertParagraph_paragraphText__insertLocation_)|Inserts a paragraph at the specified location.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#insertText_text__insertLocation_)|Inserts text into the paragraph at the specified location.|
||[leftIndent](/javascript/api/word/word.paragraph#leftIndent)|Gets or sets the left indent value, in points, for the paragraph.|
||[lineSpacing](/javascript/api/word/word.paragraph#lineSpacing)|Gets or sets the line spacing, in points, for the specified paragraph.|
||[lineUnitAfter](/javascript/api/word/word.paragraph#lineUnitAfter)|Gets or sets the amount of spacing, in grid lines, after the paragraph.|
||[lineUnitBefore](/javascript/api/word/word.paragraph#lineUnitBefore)|Gets or sets the amount of spacing, in grid lines, before the paragraph.|
||[matchCase](/javascript/api/word/word.paragraph#matchCase)||
||[matchPrefix](/javascript/api/word/word.paragraph#matchPrefix)||
||[matchSuffix](/javascript/api/word/word.paragraph#matchSuffix)||
||[matchWholeWord](/javascript/api/word/word.paragraph#matchWholeWord)||
||[matchWildcards](/javascript/api/word/word.paragraph#matchWildcards)||
||[outlineLevel](/javascript/api/word/word.paragraph#outlineLevel)|Gets or sets the outline level for the paragraph.|
||[parentContentControl](/javascript/api/word/word.paragraph#parentContentControl)|Gets the content control that contains the paragraph.|
||[rightIndent](/javascript/api/word/word.paragraph#rightIndent)|Gets or sets the right indent value, in points, for the paragraph.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| {            ignorePunct?: boolean            ignoreSpace?: boolean            matchCase?: boolean            matchPrefix?: boolean            matchSuffix?: boolean            matchWholeWord?: boolean            matchWildcards?: boolean        })](/javascript/api/word/word.paragraph#search_searchText__searchOptions__ignorePunct__ignoreSpace__matchCase__matchPrefix__matchSuffix__matchWholeWord__matchWildcards_)|Performs a search with the specified SearchOptions on the scope of the paragraph object.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.paragraph#select_selectionMode_)|Selects and navigates the Word UI to the paragraph.|
||[spaceAfter](/javascript/api/word/word.paragraph#spaceAfter)|Gets or sets the spacing, in points, after the paragraph.|
||[spaceBefore](/javascript/api/word/word.paragraph#spaceBefore)|Gets or sets the spacing, in points, before the paragraph.|
||[style](/javascript/api/word/word.paragraph#style)|Gets or sets the style name for the paragraph.|
||[text](/javascript/api/word/word.paragraph#text)|Gets the text of the paragraph.|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[items](/javascript/api/word/word.paragraphcollection#items)|Gets the loaded child items in this collection.|
|[Range](/javascript/api/word/word.range)|[clear()](/javascript/api/word/word.range#clear__)|Clears the contents of the range object.|
||[contentControls](/javascript/api/word/word.range#contentControls)|Gets the collection of content control objects in the range.|
||[delete()](/javascript/api/word/word.range#delete__)|Deletes the range and its content from the document.|
||[font](/javascript/api/word/word.range#font)|Gets the text format of the range.|
||[getHtml()](/javascript/api/word/word.range#getHtml__)|Gets an HTML representation of the range object.|
||[getOoxml()](/javascript/api/word/word.range#getOoxml__)|Gets the OOXML representation of the range object.|
||[ignorePunct](/javascript/api/word/word.range#ignorePunct)||
||[ignoreSpace](/javascript/api/word/word.range#ignoreSpace)||
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#insertBreak_breakType__insertLocation_)|Inserts a break at the specified location in the main document.|
||[insertContentControl()](/javascript/api/word/word.range#insertContentControl__)|Wraps the range object with a rich text content control.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#insertFileFromBase64_base64File__insertLocation_)|Inserts a document at the specified location.|
||[insertHtml(html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#insertHtml_html__insertLocation_)|Inserts HTML at the specified location.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#insertOoxml_ooxml__insertLocation_)|Inserts OOXML at the specified location.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#insertParagraph_paragraphText__insertLocation_)|Inserts a paragraph at the specified location.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#insertText_text__insertLocation_)|Inserts text at the specified location.|
||[matchCase](/javascript/api/word/word.range#matchCase)||
||[matchPrefix](/javascript/api/word/word.range#matchPrefix)||
||[matchSuffix](/javascript/api/word/word.range#matchSuffix)||
||[matchWholeWord](/javascript/api/word/word.range#matchWholeWord)||
||[matchWildcards](/javascript/api/word/word.range#matchWildcards)||
||[paragraphs](/javascript/api/word/word.range#paragraphs)|Gets the collection of paragraph objects in the range.|
||[parentContentControl](/javascript/api/word/word.range#parentContentControl)|Gets the content control that contains the range.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| {            ignorePunct?: boolean            ignoreSpace?: boolean            matchCase?: boolean            matchPrefix?: boolean            matchSuffix?: boolean            matchWholeWord?: boolean            matchWildcards?: boolean        })](/javascript/api/word/word.range#search_searchText__searchOptions__ignorePunct__ignoreSpace__matchCase__matchPrefix__matchSuffix__matchWholeWord__matchWildcards_)|Performs a search with the specified SearchOptions on the scope of the range object.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.range#select_selectionMode_)|Selects and navigates the Word UI to the range.|
||[style](/javascript/api/word/word.range#style)|Gets or sets the style name for the range.|
||[text](/javascript/api/word/word.range#text)|Gets the text of the range.|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[items](/javascript/api/word/word.rangecollection#items)|Gets the loaded child items in this collection.|
|[SearchOptions](/javascript/api/word/word.searchoptions)|[ignorePunct](/javascript/api/word/word.searchoptions#ignorePunct)|Gets or sets a value that indicates whether to ignore all punctuation characters between words.|
||[ignoreSpace](/javascript/api/word/word.searchoptions#ignoreSpace)|Gets or sets a value that indicates whether to ignore all whitespace between words.|
||[matchCase](/javascript/api/word/word.searchoptions#matchCase)|Gets or sets a value that indicates whether to perform a case sensitive search.|
||[matchPrefix](/javascript/api/word/word.searchoptions#matchPrefix)|Gets or sets a value that indicates whether to match words that begin with the search string.|
||[matchSuffix](/javascript/api/word/word.searchoptions#matchSuffix)|Gets or sets a value that indicates whether to match words that end with the search string.|
||[matchWholeWord](/javascript/api/word/word.searchoptions#matchWholeWord)|Gets or sets a value that indicates whether to find operation only entire words, not text that is part of a larger word.|
||[matchWildcards](/javascript/api/word/word.searchoptions#matchWildcards)|Gets or sets a value that indicates whether the search will be performed using special search operators.|
|[Section](/javascript/api/word/word.section)|[body](/javascript/api/word/word.section#body)|Gets the body object of the section.|
||[getFooter(type: Word.HeaderFooterType)](/javascript/api/word/word.section#getFooter_type_)|Gets one of the section's footers.|
||[getHeader(type: Word.HeaderFooterType)](/javascript/api/word/word.section#getHeader_type_)|Gets one of the section's headers.|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[items](/javascript/api/word/word.sectioncollection#items)|Gets the loaded child items in this collection.|

## See also

- [Word JavaScript API Reference Documentation](/javascript/api/word)
- [Word JavaScript API requirement sets](word-api-requirement-sets.md)
