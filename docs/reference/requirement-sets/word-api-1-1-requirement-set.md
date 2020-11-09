---
title: Word JavaScript API requirement set 1.1
description: 'Details about the WordApi 1.1 requirement set'
ms.date: 11/09/2020
ms.prod: word
localization_priority: Normal
---

# What's new in Word JavaScript API 1.1

WordApi 1.1 is the first requirement set of the Word JavaScript API. It's the only Word API requirement set supported by Word 2016.

## API list

The following table lists the APIs in Word JavaScript API requirement set 1.1. To view API reference documentation for all APIs supported by Word JavaScript API requirement set 1.1, see [Word APIs in requirement set 1.1](/javascript/api/word?view=word-js-1.1&preserve-view=true).

| Class | Fields | Description |
|:---|:---|:---|
|[Body](/javascript/api/word/word.body)|[clear()](/javascript/api/word/word.body#clear--)|Clears the contents of the body object.|
||[getHtml()](/javascript/api/word/word.body#gethtml--)|Gets an HTML representation of the body object.|
||[getOoxml()](/javascript/api/word/word.body#getooxml--)|Gets the OOXML (Office Open XML) representation of the body object.|
||[ignorePunct](/javascript/api/word/word.body#ignorepunct)||
||[ignoreSpace](/javascript/api/word/word.body#ignorespace)||
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#insertbreak-breaktype--insertlocation-)|Inserts a break at the specified location in the main document.|
||[insertContentControl()](/javascript/api/word/word.body#insertcontentcontrol--)|Wraps the body object with a Rich Text content control.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#insertfilefrombase64-base64file--insertlocation-)|Inserts a document into the body at the specified location.|
||[insertHtml(html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#inserthtml-html--insertlocation-)|Inserts HTML at the specified location.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#insertooxml-ooxml--insertlocation-)|Inserts OOXML at the specified location.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#insertparagraph-paragraphtext--insertlocation-)|Inserts a paragraph at the specified location.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#inserttext-text--insertlocation-)|Inserts text into the body at the specified location.|
||[matchCase](/javascript/api/word/word.body#matchcase)||
||[matchPrefix](/javascript/api/word/word.body#matchprefix)||
||[matchSuffix](/javascript/api/word/word.body#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.body#matchwholeword)||
||[matchWildcards](/javascript/api/word/word.body#matchwildcards)||
||[contentControls](/javascript/api/word/word.body#contentcontrols)|Gets the collection of rich text content control objects in the body.|
||[font](/javascript/api/word/word.body#font)|Gets the text format of the body.|
||[inlinePictures](/javascript/api/word/word.body#inlinepictures)|Gets the collection of InlinePicture objects in the body.|
||[paragraphs](/javascript/api/word/word.body#paragraphs)|Gets the collection of paragraph objects in the body.|
||[parentContentControl](/javascript/api/word/word.body#parentcontentcontrol)|Gets the content control that contains the body.|
||[text](/javascript/api/word/word.body#text)|Gets the text of the body.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| {            ignorePunct?: boolean            ignoreSpace?: boolean            matchCase?: boolean            matchPrefix?: boolean            matchSuffix?: boolean            matchWholeWord?: boolean            matchWildcards?: boolean        })](/javascript/api/word/word.body#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|Performs a search with the specified SearchOptions on the scope of the body object.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.body#select-selectionmode-)|Selects the body and navigates the Word UI to it.|
||[style](/javascript/api/word/word.body#style)|Gets or sets the style name for the body.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[appearance](/javascript/api/word/word.contentcontrol#appearance)|Gets or sets the appearance of the content control.|
||[cannotDelete](/javascript/api/word/word.contentcontrol#cannotdelete)|Gets or sets a value that indicates whether the user can delete the content control.|
||[cannotEdit](/javascript/api/word/word.contentcontrol#cannotedit)|Gets or sets a value that indicates whether the user can edit the contents of the content control.|
||[clear()](/javascript/api/word/word.contentcontrol#clear--)|Clears the contents of the content control.|
||[color](/javascript/api/word/word.contentcontrol#color)|Gets or sets the color of the content control.|
||[delete(keepContent: boolean)](/javascript/api/word/word.contentcontrol#delete-keepcontent-)|Deletes the content control and its content.|
||[getHtml()](/javascript/api/word/word.contentcontrol#gethtml--)|Gets an HTML representation of the content control object.|
||[getOoxml()](/javascript/api/word/word.contentcontrol#getooxml--)|Gets the Office Open XML (OOXML) representation of the content control object.|
||[ignorePunct](/javascript/api/word/word.contentcontrol#ignorepunct)||
||[ignoreSpace](/javascript/api/word/word.contentcontrol#ignorespace)||
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#insertbreak-breaktype--insertlocation-)|Inserts a break at the specified location in the main document.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#insertfilefrombase64-base64file--insertlocation-)|Inserts a document into the content control at the specified location.|
||[insertHtml(html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#inserthtml-html--insertlocation-)|Inserts HTML into the content control at the specified location.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#insertooxml-ooxml--insertlocation-)|Inserts OOXML into the content control at the specified location.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#insertparagraph-paragraphtext--insertlocation-)|Inserts a paragraph at the specified location.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#inserttext-text--insertlocation-)|Inserts text into the content control at the specified location.|
||[matchCase](/javascript/api/word/word.contentcontrol#matchcase)||
||[matchPrefix](/javascript/api/word/word.contentcontrol#matchprefix)||
||[matchSuffix](/javascript/api/word/word.contentcontrol#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.contentcontrol#matchwholeword)||
||[matchWildcards](/javascript/api/word/word.contentcontrol#matchwildcards)||
||[placeholderText](/javascript/api/word/word.contentcontrol#placeholdertext)|Gets or sets the placeholder text of the content control.|
||[contentControls](/javascript/api/word/word.contentcontrol#contentcontrols)|Gets the collection of content control objects in the content control.|
||[font](/javascript/api/word/word.contentcontrol#font)|Gets the text format of the content control.|
||[id](/javascript/api/word/word.contentcontrol#id)|Gets an integer that represents the content control identifier.|
||[inlinePictures](/javascript/api/word/word.contentcontrol#inlinepictures)|Gets the collection of inlinePicture objects in the content control.|
||[paragraphs](/javascript/api/word/word.contentcontrol#paragraphs)|Get the collection of paragraph objects in the content control.|
||[parentContentControl](/javascript/api/word/word.contentcontrol#parentcontentcontrol)|Gets the content control that contains the content control.|
||[text](/javascript/api/word/word.contentcontrol#text)|Gets the text of the content control.|
||[type](/javascript/api/word/word.contentcontrol#type)|Gets the content control type.|
||[removeWhenEdited](/javascript/api/word/word.contentcontrol#removewhenedited)|Gets or sets a value that indicates whether the content control is removed after it is edited.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| {            ignorePunct?: boolean            ignoreSpace?: boolean            matchCase?: boolean            matchPrefix?: boolean            matchSuffix?: boolean            matchWholeWord?: boolean            matchWildcards?: boolean        })](/javascript/api/word/word.contentcontrol#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|Performs a search with the specified SearchOptions on the scope of the content control object.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.contentcontrol#select-selectionmode-)|Selects the content control.|
||[style](/javascript/api/word/word.contentcontrol#style)|Gets or sets the style name for the content control.|
||[tag](/javascript/api/word/word.contentcontrol#tag)|Gets or sets a tag to identify a content control.|
||[title](/javascript/api/word/word.contentcontrol#title)|Gets or sets the title for a content control.|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getById(id: number)](/javascript/api/word/word.contentcontrolcollection#getbyid-id-)|Gets a content control by its identifier.|
||[getByTag(tag: string)](/javascript/api/word/word.contentcontrolcollection#getbytag-tag-)|Gets the content controls that have the specified tag.|
||[getByTitle(title: string)](/javascript/api/word/word.contentcontrolcollection#getbytitle-title-)|Gets the content controls that have the specified title.|
||[getItem(index: number)](/javascript/api/word/word.contentcontrolcollection#getitem-index-)|Gets a content control by its index in the collection.|
||[items](/javascript/api/word/word.contentcontrolcollection#items)|Gets the loaded child items in this collection.|
|[Document](/javascript/api/word/word.document)|[getSelection()](/javascript/api/word/word.document#getselection--)|Gets the current selection of the document.|
||[body](/javascript/api/word/word.document#body)|Gets the body object of the document.|
||[contentControls](/javascript/api/word/word.document#contentcontrols)|Gets the collection of content control objects in the document.|
||[saved](/javascript/api/word/word.document#saved)|Indicates whether the changes in the document have been saved.|
||[sections](/javascript/api/word/word.document#sections)|Gets the collection of section objects in the document.|
||[save()](/javascript/api/word/word.document#save--)|Saves the document.|
|[Font](/javascript/api/word/word.font)|[bold](/javascript/api/word/word.font#bold)|Gets or sets a value that indicates whether the font is bold.|
||[color](/javascript/api/word/word.font#color)|Gets or sets the color for the specified font.|
||[doubleStrikeThrough](/javascript/api/word/word.font#doublestrikethrough)|Gets or sets a value that indicates whether the font has a double strikethrough.|
||[highlightColor](/javascript/api/word/word.font#highlightcolor)|Gets or sets the highlight color.|
||[italic](/javascript/api/word/word.font#italic)|Gets or sets a value that indicates whether the font is italicized.|
||[name](/javascript/api/word/word.font#name)|Gets or sets a value that represents the name of the font.|
||[size](/javascript/api/word/word.font#size)|Gets or sets a value that represents the font size in points.|
||[strikeThrough](/javascript/api/word/word.font#strikethrough)|Gets or sets a value that indicates whether the font has a strikethrough.|
||[subscript](/javascript/api/word/word.font#subscript)|Gets or sets a value that indicates whether the font is a subscript.|
||[superscript](/javascript/api/word/word.font#superscript)|Gets or sets a value that indicates whether the font is a superscript.|
||[underline](/javascript/api/word/word.font#underline)|Gets or sets a value that indicates the font's underline type.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[altTextDescription](/javascript/api/word/word.inlinepicture#alttextdescription)|Gets or sets a string that represents the alternative text associated with the inline image.|
||[altTextTitle](/javascript/api/word/word.inlinepicture#alttexttitle)|Gets or sets a string that contains the title for the inline image.|
||[getBase64ImageSrc()](/javascript/api/word/word.inlinepicture#getbase64imagesrc--)|Gets the base64 encoded string representation of the inline image.|
||[height](/javascript/api/word/word.inlinepicture#height)|Gets or sets a number that describes the height of the inline image.|
||[hyperlink](/javascript/api/word/word.inlinepicture#hyperlink)|Gets or sets a hyperlink on the image.|
||[insertContentControl()](/javascript/api/word/word.inlinepicture#insertcontentcontrol--)|Wraps the inline picture with a rich text content control.|
||[lockAspectRatio](/javascript/api/word/word.inlinepicture#lockaspectratio)|Gets or sets a value that indicates whether the inline image retains its original proportions when you resize it.|
||[parentContentControl](/javascript/api/word/word.inlinepicture#parentcontentcontrol)|Gets the content control that contains the inline image.|
||[width](/javascript/api/word/word.inlinepicture#width)|Gets or sets a number that describes the width of the inline image.|
|[InlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|[items](/javascript/api/word/word.inlinepicturecollection#items)|Gets the loaded child items in this collection.|
|[Paragraph](/javascript/api/word/word.paragraph)|[alignment](/javascript/api/word/word.paragraph#alignment)|Gets or sets the alignment for a paragraph.|
||[clear()](/javascript/api/word/word.paragraph#clear--)|Clears the contents of the paragraph object.|
||[delete()](/javascript/api/word/word.paragraph#delete--)|Deletes the paragraph and its content from the document.|
||[firstLineIndent](/javascript/api/word/word.paragraph#firstlineindent)|Gets or sets the value, in points, for a first line or hanging indent.|
||[getHtml()](/javascript/api/word/word.paragraph#gethtml--)|Gets an HTML representation of the paragraph object.|
||[getOoxml()](/javascript/api/word/word.paragraph#getooxml--)|Gets the Office Open XML (OOXML) representation of the paragraph object.|
||[ignorePunct](/javascript/api/word/word.paragraph#ignorepunct)||
||[ignoreSpace](/javascript/api/word/word.paragraph#ignorespace)||
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#insertbreak-breaktype--insertlocation-)|Inserts a break at the specified location in the main document.|
||[insertContentControl()](/javascript/api/word/word.paragraph#insertcontentcontrol--)|Wraps the paragraph object with a rich text content control.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#insertfilefrombase64-base64file--insertlocation-)|Inserts a document into the paragraph at the specified location.|
||[insertHtml(html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#inserthtml-html--insertlocation-)|Inserts HTML into the paragraph at the specified location.|
||[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#insertinlinepicturefrombase64-base64encodedimage--insertlocation-)|Inserts a picture into the paragraph at the specified location.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#insertooxml-ooxml--insertlocation-)|Inserts OOXML into the paragraph at the specified location.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#insertparagraph-paragraphtext--insertlocation-)|Inserts a paragraph at the specified location.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#inserttext-text--insertlocation-)|Inserts text into the paragraph at the specified location.|
||[leftIndent](/javascript/api/word/word.paragraph#leftindent)|Gets or sets the left indent value, in points, for the paragraph.|
||[lineSpacing](/javascript/api/word/word.paragraph#linespacing)|Gets or sets the line spacing, in points, for the specified paragraph.|
||[lineUnitAfter](/javascript/api/word/word.paragraph#lineunitafter)|Gets or sets the amount of spacing, in grid lines, after the paragraph.|
||[lineUnitBefore](/javascript/api/word/word.paragraph#lineunitbefore)|Gets or sets the amount of spacing, in grid lines, before the paragraph.|
||[matchCase](/javascript/api/word/word.paragraph#matchcase)||
||[matchPrefix](/javascript/api/word/word.paragraph#matchprefix)||
||[matchSuffix](/javascript/api/word/word.paragraph#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.paragraph#matchwholeword)||
||[matchWildcards](/javascript/api/word/word.paragraph#matchwildcards)||
||[outlineLevel](/javascript/api/word/word.paragraph#outlinelevel)|Gets or sets the outline level for the paragraph.|
||[contentControls](/javascript/api/word/word.paragraph#contentcontrols)|Gets the collection of content control objects in the paragraph.|
||[font](/javascript/api/word/word.paragraph#font)|Gets the text format of the paragraph.|
||[inlinePictures](/javascript/api/word/word.paragraph#inlinepictures)|Gets the collection of InlinePicture objects in the paragraph.|
||[parentContentControl](/javascript/api/word/word.paragraph#parentcontentcontrol)|Gets the content control that contains the paragraph.|
||[text](/javascript/api/word/word.paragraph#text)|Gets the text of the paragraph.|
||[rightIndent](/javascript/api/word/word.paragraph#rightindent)|Gets or sets the right indent value, in points, for the paragraph.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| {            ignorePunct?: boolean            ignoreSpace?: boolean            matchCase?: boolean            matchPrefix?: boolean            matchSuffix?: boolean            matchWholeWord?: boolean            matchWildcards?: boolean        })](/javascript/api/word/word.paragraph#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|Performs a search with the specified SearchOptions on the scope of the paragraph object.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.paragraph#select-selectionmode-)|Selects and navigates the Word UI to the paragraph.|
||[spaceAfter](/javascript/api/word/word.paragraph#spaceafter)|Gets or sets the spacing, in points, after the paragraph.|
||[spaceBefore](/javascript/api/word/word.paragraph#spacebefore)|Gets or sets the spacing, in points, before the paragraph.|
||[style](/javascript/api/word/word.paragraph#style)|Gets or sets the style name for the paragraph.|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[items](/javascript/api/word/word.paragraphcollection#items)|Gets the loaded child items in this collection.|
|[Range](/javascript/api/word/word.range)|[clear()](/javascript/api/word/word.range#clear--)|Clears the contents of the range object.|
||[delete()](/javascript/api/word/word.range#delete--)|Deletes the range and its content from the document.|
||[getHtml()](/javascript/api/word/word.range#gethtml--)|Gets an HTML representation of the range object.|
||[getOoxml()](/javascript/api/word/word.range#getooxml--)|Gets the OOXML representation of the range object.|
||[ignorePunct](/javascript/api/word/word.range#ignorepunct)||
||[ignoreSpace](/javascript/api/word/word.range#ignorespace)||
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#insertbreak-breaktype--insertlocation-)|Inserts a break at the specified location in the main document.|
||[insertContentControl()](/javascript/api/word/word.range#insertcontentcontrol--)|Wraps the range object with a rich text content control.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#insertfilefrombase64-base64file--insertlocation-)|Inserts a document at the specified location.|
||[insertHtml(html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#inserthtml-html--insertlocation-)|Inserts HTML at the specified location.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#insertooxml-ooxml--insertlocation-)|Inserts OOXML at the specified location.|
||[insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#insertparagraph-paragraphtext--insertlocation-)|Inserts a paragraph at the specified location.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#inserttext-text--insertlocation-)|Inserts text at the specified location.|
||[matchCase](/javascript/api/word/word.range#matchcase)||
||[matchPrefix](/javascript/api/word/word.range#matchprefix)||
||[matchSuffix](/javascript/api/word/word.range#matchsuffix)||
||[matchWholeWord](/javascript/api/word/word.range#matchwholeword)||
||[matchWildcards](/javascript/api/word/word.range#matchwildcards)||
||[contentControls](/javascript/api/word/word.range#contentcontrols)|Gets the collection of content control objects in the range.|
||[font](/javascript/api/word/word.range#font)|Gets the text format of the range.|
||[paragraphs](/javascript/api/word/word.range#paragraphs)|Gets the collection of paragraph objects in the range.|
||[parentContentControl](/javascript/api/word/word.range#parentcontentcontrol)|Gets the content control that contains the range.|
||[text](/javascript/api/word/word.range#text)|Gets the text of the range.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| {            ignorePunct?: boolean            ignoreSpace?: boolean            matchCase?: boolean            matchPrefix?: boolean            matchSuffix?: boolean            matchWholeWord?: boolean            matchWildcards?: boolean        })](/javascript/api/word/word.range#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|Performs a search with the specified SearchOptions on the scope of the range object.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.range#select-selectionmode-)|Selects and navigates the Word UI to the range.|
||[style](/javascript/api/word/word.range#style)|Gets or sets the style name for the range.|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[items](/javascript/api/word/word.rangecollection#items)|Gets the loaded child items in this collection.|
|[SearchOptions](/javascript/api/word/word.searchoptions)|[ignorePunct](/javascript/api/word/word.searchoptions#ignorepunct)|Gets or sets a value that indicates whether to ignore all punctuation characters between words.|
||[ignoreSpace](/javascript/api/word/word.searchoptions#ignorespace)|Gets or sets a value that indicates whether to ignore all whitespace between words.|
||[matchCase](/javascript/api/word/word.searchoptions#matchcase)|Gets or sets a value that indicates whether to perform a case sensitive search.|
||[matchPrefix](/javascript/api/word/word.searchoptions#matchprefix)|Gets or sets a value that indicates whether to match words that begin with the search string.|
||[matchSuffix](/javascript/api/word/word.searchoptions#matchsuffix)|Gets or sets a value that indicates whether to match words that end with the search string.|
||[matchWholeWord](/javascript/api/word/word.searchoptions#matchwholeword)|Gets or sets a value that indicates whether to find operation only entire words, not text that is part of a larger word.|
||[matchWildcards](/javascript/api/word/word.searchoptions#matchwildcards)|Gets or sets a value that indicates whether the search will be performed using special search operators.|
|[Section](/javascript/api/word/word.section)|[getFooter(type: Word.HeaderFooterType)](/javascript/api/word/word.section#getfooter-type-)|Gets one of the section's footers.|
||[getHeader(type: Word.HeaderFooterType)](/javascript/api/word/word.section#getheader-type-)|Gets one of the section's headers.|
||[body](/javascript/api/word/word.section#body)|Gets the body object of the section.|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[items](/javascript/api/word/word.sectioncollection#items)|Gets the loaded child items in this collection.|

## See also

- [Word JavaScript API Reference Documentation](/javascript/api/word)
- [Word JavaScript API requirement sets](word-api-requirement-sets.md)
