**Resource name:** [application](application.md)

**What's new:** Method **methodPlusLink** returning **[Document](document.md)**

**Description:** Creates a new document by using a base64 encoded .docx file.

**Available in requirement set:** WordApiDesktop, 1.3

_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=application-createDocument)_


**Resource name:** [body](body.md)

**What's new:** Relationship **lists** of type **[ListCollection](listcollection.md)**

**Description:** Gets the collection of list objects in the body. Read-only.

**Available in requirement set:** 1.3

_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=body-lists)_



**Resource name:** [body](body.md)
**What's new:** Relationship **parentBody** of type **[Body](body.md)**
**Description:** Gets the parent body of the body. For example, a table cell body's parent body could be a header. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=body-parentBody)_

**Resource name:** [body](body.md)
**What's new:** Relationship **tables** of type **[TableCollection](tablecollection.md)**
**Description:** Gets the collection of table objects in the body. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=body-tables)_

**Resource name:** [body](body.md)
**What's new:** Relationship **type** of type **[BodyType](bodytype.md)**
**Description:** Gets the type of the body. The type can be 'MainDoc', 'Section', 'Header', 'Footer', or 'TableCell'. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=body-type)_

**Resource name:** [body](body.md)
**What's new:** Method **methodPlusLink** returning **[Range](range.md)**
**Description:** Gets the whole body, or the starting or ending point of the body, as a range.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=body-getRange)_

**Resource name:** [body](body.md)
**What's new:** Method **methodPlusLink** returning **[Table](table.md)**
**Description:** Inserts a table with the specified number of rows and columns. The insertLocation value can be 'Start' or 'End'.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=body-insertTable)_

**Resource name:** [contentControl](contentcontrol.md)
**What's new:** Relationship **lists** of type **[ListCollection](listcollection.md)**
**Description:** Gets the collection of list objects in the content control. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=contentControl-lists)_

**Resource name:** [contentControl](contentcontrol.md)
**What's new:** Relationship **parentTable** of type **[Table](table.md)**
**Description:** Gets the table that contains the content control. Returns null if it is not contained in a table. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=contentControl-parentTable)_

**Resource name:** [contentControl](contentcontrol.md)
**What's new:** Relationship **parentTableCell** of type **[TableCell](tablecell.md)**
**Description:** Gets the table cell that contains the content control. Returns null if it is not contained in a table cell. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=contentControl-parentTableCell)_

**Resource name:** [contentControl](contentcontrol.md)
**What's new:** Relationship **subtype** of type **[ContentControlType](contentcontroltype.md)**
**Description:** Gets the content control subtype. The subtype can be 'RichTextInline', 'RichTextParagraphs', 'RichTextTableCell', 'RichTextTableRow' and 'RichTextTable' for rich text content controls. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=contentControl-subtype)_

**Resource name:** [contentControl](contentcontrol.md)
**What's new:** Relationship **tables** of type **[TableCollection](tablecollection.md)**
**Description:** Gets the collection of table objects in the content control. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=contentControl-tables)_

**Resource name:** [contentControl](contentcontrol.md)
**What's new:** Method **methodPlusLink** returning **[Range](range.md)**
**Description:** Gets the whole content control, or the starting or ending point of the content control, as a range.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=contentControl-getRange)_

**Resource name:** [contentControl](contentcontrol.md)
**What's new:** Method **methodPlusLink** returning **[RangeCollection](rangecollection.md)**
**Description:** Gets the text ranges in the content control by using punctuation marks andor space character.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=contentControl-getTextRanges)_

**Resource name:** [contentControl](contentcontrol.md)
**What's new:** Method **methodPlusLink** returning **[Table](table.md)**
**Description:** Inserts a table with the specified number of rows and columns into, or next to, a content control. The insertLocation value can be 'Start', 'End', 'Before' or 'After'.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=contentControl-insertTable)_

**Resource name:** [contentControl](contentcontrol.md)
**What's new:** Method **methodPlusLink** returning **[RangeCollection](rangecollection.md)**
**Description:** Splits the content control into child ranges by using delimiters.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=contentControl-split)_

**Resource name:** [contentControlCollection](contentcontrolcollection.md)
**What's new:** Method **methodPlusLink** returning **[ContentControlCollection](contentcontrolcollection.md)**
**Description:** Gets the content controls that have the specified types andor subtypes.
**Available in requirement set:** WordApiDesktop, 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=contentControlCollection-getByTypes)_

**Resource name:** [document](document.md)
**What's new:** Method **methodPlusLink** returning **void**
**Description:** Open the document.
**Available in requirement set:** WordApiDesktop, 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=document-open)_

**Resource name:** [font](font.md)
**What's new:** Property **doubleStrikeThrough** of type **bool**
**Description:** Gets or sets a value that indicates whether the font has a double strike through. True if the font is formatted as double strikethrough text, otherwise, false.
**Available in requirement set:** WordApiDesktop, 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=font-doubleStrikeThrough)_

**Resource name:** [inlinePicture](inlinepicture.md)
**What's new:** Relationship **imageFormat** of type **[ImageFormat](imageformat.md)**
**Description:** Gets the format of the inline image. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=inlinePicture-imageFormat)_

**Resource name:** [inlinePicture](inlinepicture.md)
**What's new:** Relationship **next** of type **[InlinePicture](inlinepicture.md)**
**Description:** Gets the next inline image. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=inlinePicture-next)_

**Resource name:** [inlinePicture](inlinepicture.md)
**What's new:** Relationship **parentTable** of type **[Table](table.md)**
**Description:** Gets the table that contains the inline image. Returns null if it is not contained in a table. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=inlinePicture-parentTable)_

**Resource name:** [inlinePicture](inlinepicture.md)
**What's new:** Relationship **parentTableCell** of type **[TableCell](tablecell.md)**
**Description:** Gets the table cell that contains the inline image. Returns null if it is not contained in a table cell. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=inlinePicture-parentTableCell)_

**Resource name:** [inlinePicture](inlinepicture.md)
**What's new:** Method **methodPlusLink** returning **[Range](range.md)**
**Description:** Gets the picture, or the starting or ending point of the picture, as a range.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=inlinePicture-getRange)_

**Resource name:** [inlinePictureCollection](inlinepicturecollection.md)
**What's new:** Relationship **first** of type **[InlinePicture](inlinepicture.md)**
**Description:** Gets the first inline image in this collection. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=inlinePictureCollection-first)_

**Resource name:** [list](list.md)
**What's new:** Property **id** of type **int**
**Description:** Gets the list's id. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=list-id)_

**Resource name:** [list](list.md)
**What's new:** Relationship **paragraphs** of type **[ParagraphCollection](paragraphcollection.md)**
**Description:** A collection containing the paragraphs in this list. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=list-paragraphs)_

**Resource name:** [list](list.md)
**What's new:** Method **methodPlusLink** returning **[Paragraph](paragraph.md)**
**Description:** Inserts a paragraph at the specified location. The insertLocation value can be 'Start', 'End', 'Before' or 'After'.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=list-insertParagraph)_

**Resource name:** [listCollection](listcollection.md)
**What's new:** Property **items** of type **[List[]](list.md)**
**Description:** A collection of list objects. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=listCollection-items)_

**Resource name:** [listCollection](listcollection.md)
**What's new:** Relationship **first** of type **[List](list.md)**
**Description:** Gets the first list in this collection. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=listCollection-first)_

**Resource name:** [listCollection](listcollection.md)
**What's new:** Method **methodPlusLink** returning **[List](list.md)**
**Description:** Gets a list by its identifier.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=listCollection-getById)_

**Resource name:** [listCollection](listcollection.md)
**What's new:** Method **methodPlusLink** returning **[List](list.md)**
**Description:** Gets a list object by its index in the collection.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=listCollection-getItem)_

**Resource name:** [paragraph](paragraph.md)
**What's new:** Property **listLevel** of type **int**
**Description:** Gets or sets the list level of the paragraph.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=paragraph-listLevel)_

**Resource name:** [paragraph](paragraph.md)
**What's new:** Property **outlineLevel** of type **int**
**Description:** Gets or sets the outline level for the paragraph.
**Available in requirement set:** WordApiDesktop, 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=paragraph-outlineLevel)_

**Resource name:** [paragraph](paragraph.md)
**What's new:** Property **tableNestingLevel** of type **int**
**Description:** Gets the level of the paragraph's table. It returns 0 if the paragraph is not in a table. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=paragraph-tableNestingLevel)_

**Resource name:** [paragraph](paragraph.md)
**What's new:** Relationship **list** of type **[List](list.md)**
**Description:** Gets the List to which this paragraph belongs. Returns null if the paragraph is not in a list. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=paragraph-list)_

**Resource name:** [paragraph](paragraph.md)
**What's new:** Relationship **next** of type **[Paragraph](paragraph.md)**
**Description:** Gets the next paragraph. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=paragraph-next)_

**Resource name:** [paragraph](paragraph.md)
**What's new:** Relationship **parentBody** of type **[Body](body.md)**
**Description:** Gets the parent body of the paragraph. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=paragraph-parentBody)_

**Resource name:** [paragraph](paragraph.md)
**What's new:** Relationship **parentTable** of type **[Table](table.md)**
**Description:** Gets the table that contains the paragraph. Returns null if it is not contained in a table. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=paragraph-parentTable)_

**Resource name:** [paragraph](paragraph.md)
**What's new:** Relationship **parentTableCell** of type **[TableCell](tablecell.md)**
**Description:** Gets the table cell that contains the paragraph. Returns null if it is not contained in a table cell. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=paragraph-parentTableCell)_

**Resource name:** [paragraph](paragraph.md)
**What's new:** Relationship **previous** of type **[Paragraph](paragraph.md)**
**Description:** Gets the previous paragraph. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=paragraph-previous)_

**Resource name:** [paragraph](paragraph.md)
**What's new:** Method **methodPlusLink** returning **[Range](range.md)**
**Description:** Gets the whole paragraph, or the starting or ending point of the paragraph, as a range.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=paragraph-getRange)_

**Resource name:** [paragraph](paragraph.md)
**What's new:** Method **methodPlusLink** returning **[RangeCollection](rangecollection.md)**
**Description:** Gets the text ranges in the paragraph by using punctuation marks andor space character.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=paragraph-getTextRanges)_

**Resource name:** [paragraph](paragraph.md)
**What's new:** Method **methodPlusLink** returning **[Table](table.md)**
**Description:** Inserts a table with the specified number of rows and columns. The insertLocation value can be 'Before' or 'After'.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=paragraph-insertTable)_

**Resource name:** [paragraph](paragraph.md)
**What's new:** Method **methodPlusLink** returning **[RangeCollection](rangecollection.md)**
**Description:** Splits the paragraph into child ranges by using delimiters.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=paragraph-split)_

**Resource name:** [paragraphCollection](paragraphcollection.md)
**What's new:** Relationship **first** of type **[Paragraph](paragraph.md)**
**Description:** Gets the first paragraph in this collection. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=paragraphCollection-first)_

**Resource name:** [paragraphCollection](paragraphcollection.md)
**What's new:** Relationship **last** of type **[Paragraph](paragraph.md)**
**Description:** Gets the last paragraph in this collection. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=paragraphCollection-last)_

**Resource name:** [range](range.md)
**What's new:** Property **hyperlink** of type **string**
**Description:** Gets the first hyperlink in the range, or sets a hyperlink on the range. Existing hyperlinks in this range are deleted when you set a new hyperlink.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=range-hyperlink)_

**Resource name:** [range](range.md)
**What's new:** Property **isEmpty** of type **bool**
**Description:** Checks whether the range length is zero. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=range-isEmpty)_

**Resource name:** [range](range.md)
**What's new:** Relationship **lists** of type **[ListCollection](listcollection.md)**
**Description:** Gets the collection of list objects in the range. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=range-lists)_

**Resource name:** [range](range.md)
**What's new:** Relationship **parentBody** of type **[Body](body.md)**
**Description:** Gets the parent body of the range. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=range-parentBody)_

**Resource name:** [range](range.md)
**What's new:** Relationship **parentTable** of type **[Table](table.md)**
**Description:** Gets the table that contains the range. Returns null if it is not contained in a table. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=range-parentTable)_

**Resource name:** [range](range.md)
**What's new:** Relationship **parentTableCell** of type **[TableCell](tablecell.md)**
**Description:** Gets the table cell that contains the range. Returns null if it is not contained in a table cell. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=range-parentTableCell)_

**Resource name:** [range](range.md)
**What's new:** Relationship **tables** of type **[TableCollection](tablecollection.md)**
**Description:** Gets the collection of table objects in the range. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=range-tables)_

**Resource name:** [range](range.md)
**What's new:** Method **methodPlusLink** returning **[LocationRelation](locationrelation.md)**
**Description:** Compares this range's location with another range's location.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=range-compareLocationWith)_

**Resource name:** [range](range.md)
**What's new:** Method **methodPlusLink** returning **void**
**Description:** Expands the range in either direction to cover another range.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=range-expandTo)_

**Resource name:** [range](range.md)
**What's new:** Method **methodPlusLink** returning **[RangeCollection](rangecollection.md)**
**Description:** Gets hyperlink child ranges within the range.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=range-getHyperlinkRanges)_

**Resource name:** [range](range.md)
**What's new:** Method **methodPlusLink** returning **[Range](range.md)**
**Description:** Gets the next text range by using punctuation marks andor space character.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=range-getNextTextRange)_

**Resource name:** [range](range.md)
**What's new:** Method **methodPlusLink** returning **[Range](range.md)**
**Description:** Clones the range, or gets the starting or ending point of the range as a new range.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=range-getRange)_

**Resource name:** [range](range.md)
**What's new:** Method **methodPlusLink** returning **[RangeCollection](rangecollection.md)**
**Description:** Gets the text child ranges in the range by using punctuation marks andor space character.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=range-getTextRanges)_

**Resource name:** [range](range.md)
**What's new:** Method **methodPlusLink** returning **[Table](table.md)**
**Description:** Inserts a table with the specified number of rows and columns. The insertLocation value can be 'Before' or 'After'.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=range-insertTable)_

**Resource name:** [range](range.md)
**What's new:** Method **methodPlusLink** returning **void**
**Description:** Shrinks the range to the intersection of the range with another range.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=range-intersectWith)_

**Resource name:** [range](range.md)
**What's new:** Method **methodPlusLink** returning **[RangeCollection](rangecollection.md)**
**Description:** Splits the range into child ranges by using delimiters.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=range-split)_

**Resource name:** [rangeCollection](rangecollection.md)
**What's new:** Property **items** of type **[Range[]](range.md)**
**Description:** A collection of range objects. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=rangeCollection-items)_

**Resource name:** [rangeCollection](rangecollection.md)
**What's new:** Relationship **first** of type **[Range](range.md)**
**Description:** Gets the first range in this collection. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=rangeCollection-first)_

**Resource name:** [rangeCollection](rangecollection.md)
**What's new:** Method **methodPlusLink** returning **[Range](range.md)**
**Description:** Gets a range object by its index in the collection.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=rangeCollection-getItem)_

**Resource name:** [searchResultCollection](searchresultcollection.md)
**What's new:** Relationship **first** of type **[Range](range.md)**
**Description:** Gets the first searched result in this collection. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=searchResultCollection-first)_

**Resource name:** [section](section.md)
**What's new:** Relationship **next** of type **[Section](section.md)**
**Description:** Gets the next section. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=section-next)_

**Resource name:** [sectionCollection](sectioncollection.md)
**What's new:** Relationship **first** of type **[Section](section.md)**
**Description:** Gets the first section in this collection. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=sectionCollection-first)_

**Resource name:** [table](table.md)
**What's new:** Property **headerRowCount** of type **int**
**Description:** Gets and sets the number of header rows.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-headerRowCount)_

**Resource name:** [table](table.md)
**What's new:** Property **isUniform** of type **bool**
**Description:** Indicates whether all of the table rows are uniform. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-isUniform)_

**Resource name:** [table](table.md)
**What's new:** Property **nestingLevel** of type **int**
**Description:** Gets the nesting level of the table. Top-level tables have level 1. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-nestingLevel)_

**Resource name:** [table](table.md)
**What's new:** Property **rowCount** of type **int**
**Description:** Gets the number of rows in the table. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-rowCount)_

**Resource name:** [table](table.md)
**What's new:** Property **shadingColor** of type **string**
**Description:** Gets and sets the shading color.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-shadingColor)_

**Resource name:** [table](table.md)
**What's new:** Property **style** of type **string**
**Description:** Gets and sets the name of the table style.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-style)_

**Resource name:** [table](table.md)
**What's new:** Property **styleBandedColumns** of type **bool**
**Description:** Gets and sets whether the table has banded columns.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-styleBandedColumns)_

**Resource name:** [table](table.md)
**What's new:** Property **styleBandedRows** of type **bool**
**Description:** Gets and sets whether the table has banded rows.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-styleBandedRows)_

**Resource name:** [table](table.md)
**What's new:** Property **styleFirstColumn** of type **bool**
**Description:** Gets and sets whether the table has a first column with a special style.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-styleFirstColumn)_

**Resource name:** [table](table.md)
**What's new:** Property **styleLastColumn** of type **bool**
**Description:** Gets and sets whether the table has a last column with a special style.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-styleLastColumn)_

**Resource name:** [table](table.md)
**What's new:** Property **styleTotalRow** of type **bool**
**Description:** Gets and sets whether the table has a total (last) row with a special style.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-styleTotalRow)_

**Resource name:** [table](table.md)
**What's new:** Property **values** of type **string**
**Description:** Gets and sets the text values in the table, as a 2D Javascript array.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-values)_

**Resource name:** [table](table.md)
**What's new:** Relationship **cellPaddingBottom** of type **[float](float.md)**
**Description:** Gets and sets the default bottom cell padding in points.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-cellPaddingBottom)_

**Resource name:** [table](table.md)
**What's new:** Relationship **cellPaddingLeft** of type **[float](float.md)**
**Description:** Gets and sets the default left cell padding in points.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-cellPaddingLeft)_

**Resource name:** [table](table.md)
**What's new:** Relationship **cellPaddingRight** of type **[float](float.md)**
**Description:** Gets and sets the default right cell padding in points.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-cellPaddingRight)_

**Resource name:** [table](table.md)
**What's new:** Relationship **cellPaddingTop** of type **[float](float.md)**
**Description:** Gets and sets the default top cell padding in points.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-cellPaddingTop)_

**Resource name:** [table](table.md)
**What's new:** Relationship **font** of type **[Font](font.md)**
**Description:** Gets the font. Use this to get and set font name, size, color, and other properties. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-font)_

**Resource name:** [table](table.md)
**What's new:** Relationship **height** of type **[float](float.md)**
**Description:** Gets the height of the table in points. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-height)_

**Resource name:** [table](table.md)
**What's new:** Relationship **next** of type **[Table](table.md)**
**Description:** Gets the next table. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-next)_

**Resource name:** [table](table.md)
**What's new:** Relationship **paragraphAfter** of type **[Paragraph](paragraph.md)**
**Description:** Gets the paragraph after the table. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-paragraphAfter)_

**Resource name:** [table](table.md)
**What's new:** Relationship **paragraphBefore** of type **[Paragraph](paragraph.md)**
**Description:** Gets the paragraph before the table. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-paragraphBefore)_

**Resource name:** [table](table.md)
**What's new:** Relationship **parentContentControl** of type **[ContentControl](contentcontrol.md)**
**Description:** Gets the content control that contains the table. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-parentContentControl)_

**Resource name:** [table](table.md)
**What's new:** Relationship **parentTable** of type **[Table](table.md)**
**Description:** Gets the table that contains this table. Returns null if it is not contained in a table. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-parentTable)_

**Resource name:** [table](table.md)
**What's new:** Relationship **parentTableCell** of type **[TableCell](tablecell.md)**
**Description:** Gets the table cell that contains this table. Returns null if it is not contained in a table cell. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-parentTableCell)_

**Resource name:** [table](table.md)
**What's new:** Relationship **rows** of type **[TableRowCollection](tablerowcollection.md)**
**Description:** Gets all of the table rows. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-rows)_

**Resource name:** [table](table.md)
**What's new:** Relationship **tables** of type **[TableCollection](tablecollection.md)**
**Description:** Gets the child tables nested one level deeper. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-tables)_

**Resource name:** [table](table.md)
**What's new:** Relationship **verticalAlignment** of type **[VerticalAlignment](verticalalignment.md)**
**Description:** Gets and sets the vertical alignment of every cell in the table.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-verticalAlignment)_

**Resource name:** [table](table.md)
**What's new:** Relationship **width** of type **[float](float.md)**
**Description:** Gets and sets the width of the table in points.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-width)_

**Resource name:** [table](table.md)
**What's new:** Method **methodPlusLink** returning **void**
**Description:** Adds columns to the start or end of the table, using the first or last existing column as a template. This is applicable to uniform tables. The string values, if specified, are set in the newly inserted rows.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-addColumns)_

**Resource name:** [table](table.md)
**What's new:** Method **methodPlusLink** returning **void**
**Description:** Adds rows to the start or end of the table, using the first or last existing row as a template. The string values, if specified, are set in the newly inserted rows.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-addRows)_

**Resource name:** [table](table.md)
**What's new:** Method **methodPlusLink** returning **void**
**Description:** Autofits the table columns to the width of their contents.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-autoFitContents)_

**Resource name:** [table](table.md)
**What's new:** Method **methodPlusLink** returning **void**
**Description:** Autofits the table columns to the width of the window.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-autoFitWindow)_

**Resource name:** [table](table.md)
**What's new:** Method **methodPlusLink** returning **void**
**Description:** Clears the contents of the table.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-clear)_

**Resource name:** [table](table.md)
**What's new:** Method **methodPlusLink** returning **void**
**Description:** Deletes the entire table.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-delete)_

**Resource name:** [table](table.md)
**What's new:** Method **methodPlusLink** returning **void**
**Description:** Deletes specific columns. This is applicable to uniform tables.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-deleteColumns)_

**Resource name:** [table](table.md)
**What's new:** Method **methodPlusLink** returning **void**
**Description:** Deletes specific rows.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-deleteRows)_

**Resource name:** [table](table.md)
**What's new:** Method **methodPlusLink** returning **void**
**Description:** Distributes the column widths evenly.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-distributeColumns)_

**Resource name:** [table](table.md)
**What's new:** Method **methodPlusLink** returning **void**
**Description:** Distributes the row heights evenly.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-distributeRows)_

**Resource name:** [table](table.md)
**What's new:** Method **methodPlusLink** returning **[TableBorderStyle](tableborderstyle.md)**
**Description:** Gets the border style for the specified border.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-getBorderStyle)_

**Resource name:** [table](table.md)
**What's new:** Method **methodPlusLink** returning **[TableCell](tablecell.md)**
**Description:** Gets the table cell at a specified row and column.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-getCell)_

**Resource name:** [table](table.md)
**What's new:** Method **methodPlusLink** returning **[Range](range.md)**
**Description:** Gets the range that contains this table, or the range at the start or end of the table.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-getRange)_

**Resource name:** [table](table.md)
**What's new:** Method **methodPlusLink** returning **[ContentControl](contentcontrol.md)**
**Description:** Inserts a content control on the table.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-insertContentControl)_

**Resource name:** [table](table.md)
**What's new:** Method **methodPlusLink** returning **[Paragraph](paragraph.md)**
**Description:** Inserts a paragraph at the specified location. The insertLocation value can be 'Before' or 'After'.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-insertParagraph)_

**Resource name:** [table](table.md)
**What's new:** Method **methodPlusLink** returning **[Table](table.md)**
**Description:** Inserts a table with the specified number of rows and columns. The insertLocation value can be 'Before' or 'After'.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-insertTable)_

**Resource name:** [table](table.md)
**What's new:** Method **methodPlusLink** returning **[TableCell](tablecell.md)**
**Description:** Merges the cells bounded inclusively by a first and last cell.
**Available in requirement set:** WordApiDesktop, 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-mergeCells)_

**Resource name:** [table](table.md)
**What's new:** Method **methodPlusLink** returning **[SearchResultCollection](searchresultcollection.md)**
**Description:** Performs a search with the specified searchOptions on the scope of the table object. The search results are a collection of range objects.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-search)_

**Resource name:** [table](table.md)
**What's new:** Method **methodPlusLink** returning **void**
**Description:** Selects the table, or the position at the start or end of the table, and navigates the Word UI to it.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-select)_

**Resource name:** [tableBorderStyle](tableborderstyle.md)
**What's new:** Property **color** of type **string**
**Description:** Gets or sets the table border color, as a hex value or name.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableBorderStyle-color)_

**Resource name:** [tableBorderStyle](tableborderstyle.md)
**What's new:** Relationship **type** of type **[BorderType](bordertype.md)**
**Description:** Gets or sets the type of the table border style.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableBorderStyle-type)_

**Resource name:** [tableBorderStyle](tableborderstyle.md)
**What's new:** Relationship **width** of type **[float](float.md)**
**Description:** Gets or sets the width, in points, of the table border style.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableBorderStyle-width)_

**Resource name:** [tableCell](tablecell.md)
**What's new:** Property **cellIndex** of type **int**
**Description:** Gets the index of the cell in its row. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCell-cellIndex)_

**Resource name:** [tableCell](tablecell.md)
**What's new:** Property **rowIndex** of type **int**
**Description:** Gets the index of the cell's row in the table. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCell-rowIndex)_

**Resource name:** [tableCell](tablecell.md)
**What's new:** Property **shadingColor** of type **string**
**Description:** Gets or sets the shading color of the cell. Color is specified in "#RRGGBB" format or by using the color name.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCell-shadingColor)_

**Resource name:** [tableCell](tablecell.md)
**What's new:** Property **value** of type **string**
**Description:** Gets and sets the text of the cell.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCell-value)_

**Resource name:** [tableCell](tablecell.md)
**What's new:** Relationship **body** of type **[Body](body.md)**
**Description:** Gets the body object of the cell. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCell-body)_

**Resource name:** [tableCell](tablecell.md)
**What's new:** Relationship **cellPaddingBottom** of type **[float](float.md)**
**Description:** Gets and sets the bottom padding of the cell in points.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCell-cellPaddingBottom)_

**Resource name:** [tableCell](tablecell.md)
**What's new:** Relationship **cellPaddingLeft** of type **[float](float.md)**
**Description:** Gets and sets the left padding of the cell in points.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCell-cellPaddingLeft)_

**Resource name:** [tableCell](tablecell.md)
**What's new:** Relationship **cellPaddingRight** of type **[float](float.md)**
**Description:** Gets and sets the right padding of the cell in points.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCell-cellPaddingRight)_

**Resource name:** [tableCell](tablecell.md)
**What's new:** Relationship **cellPaddingTop** of type **[float](float.md)**
**Description:** Gets and sets the top padding of the cell in points.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCell-cellPaddingTop)_

**Resource name:** [tableCell](tablecell.md)
**What's new:** Relationship **columnWidth** of type **[float](float.md)**
**Description:** Gets and sets the width of the cell's column in points. This is applicable to uniform tables.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCell-columnWidth)_

**Resource name:** [tableCell](tablecell.md)
**What's new:** Relationship **next** of type **[TableCell](tablecell.md)**
**Description:** Gets the next cell. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCell-next)_

**Resource name:** [tableCell](tablecell.md)
**What's new:** Relationship **parentRow** of type **[TableRow](tablerow.md)**
**Description:** Gets the parent row of the cell. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCell-parentRow)_

**Resource name:** [tableCell](tablecell.md)
**What's new:** Relationship **parentTable** of type **[Table](table.md)**
**Description:** Gets the parent table of the cell. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCell-parentTable)_

**Resource name:** [tableCell](tablecell.md)
**What's new:** Relationship **verticalAlignment** of type **[VerticalAlignment](verticalalignment.md)**
**Description:** Gets and sets the vertical alignment of the cell.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCell-verticalAlignment)_

**Resource name:** [tableCell](tablecell.md)
**What's new:** Relationship **width** of type **[float](float.md)**
**Description:** Gets the width of the cell in points. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCell-width)_

**Resource name:** [tableCell](tablecell.md)
**What's new:** Method **methodPlusLink** returning **void**
**Description:** Deletes the column containing this cell. This is applicable to uniform tables.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCell-deleteColumn)_

**Resource name:** [tableCell](tablecell.md)
**What's new:** Method **methodPlusLink** returning **void**
**Description:** Deletes the row containing this cell.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCell-deleteRow)_

**Resource name:** [tableCell](tablecell.md)
**What's new:** Method **methodPlusLink** returning **[TableBorderStyle](tableborderstyle.md)**
**Description:** Gets the border style for the specified border.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCell-getBorderStyle)_

**Resource name:** [tableCell](tablecell.md)
**What's new:** Method **methodPlusLink** returning **void**
**Description:** Adds columns to the left or right of the cell, using the cell's column as a template. This is applicable to uniform tables. The string values, if specified, are set in the newly inserted rows.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCell-insertColumns)_

**Resource name:** [tableCell](tablecell.md)
**What's new:** Method **methodPlusLink** returning **void**
**Description:** Inserts rows above or below the cell, using the cell's row as a template. The string values, if specified, are set in the newly inserted rows.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCell-insertRows)_

**Resource name:** [tableCell](tablecell.md)
**What's new:** Method **methodPlusLink** returning **void**
**Description:** Adds columns to the left or right of the cell, using the existing column as a template. The string values, if specified, are set in the newly inserted rows.
**Available in requirement set:** WordApiDesktop, 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCell-split)_

**Resource name:** [tableCellCollection](tablecellcollection.md)
**What's new:** Property **items** of type **[TableCell[]](tablecell.md)**
**Description:** A collection of tableCell objects. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCellCollection-items)_

**Resource name:** [tableCellCollection](tablecellcollection.md)
**What's new:** Relationship **first** of type **[TableCell](tablecell.md)**
**Description:** Gets the first table cell in this collection. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCellCollection-first)_

**Resource name:** [tableCellCollection](tablecellcollection.md)
**What's new:** Method **methodPlusLink** returning **[TableCell](tablecell.md)**
**Description:** Gets a table cell object by its index in the collection.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCellCollection-getItem)_

**Resource name:** [tableCollection](tablecollection.md)
**What's new:** Property **items** of type **[Table[]](table.md)**
**Description:** A collection of table objects. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCollection-items)_

**Resource name:** [tableCollection](tablecollection.md)
**What's new:** Relationship **first** of type **[Table](table.md)**
**Description:** Gets the first table in this collection. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCollection-first)_

**Resource name:** [tableCollection](tablecollection.md)
**What's new:** Method **methodPlusLink** returning **[Table](table.md)**
**Description:** Gets a table object by its index in the collection.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCollection-getItem)_

**Resource name:** [tableRow](tablerow.md)
**What's new:** Property **cellCount** of type **int**
**Description:** Gets the number of cells in the row. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableRow-cellCount)_

**Resource name:** [tableRow](tablerow.md)
**What's new:** Property **isHeader** of type **bool**
**Description:** Gets a value that indicates whether the row is a header row. Read-only. To set the number of header rows, use HeaderRowCount on the Table object. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableRow-isHeader)_

**Resource name:** [tableRow](tablerow.md)
**What's new:** Property **rowIndex** of type **int**
**Description:** Gets the index of the row in its parent table. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableRow-rowIndex)_

**Resource name:** [tableRow](tablerow.md)
**What's new:** Property **shadingColor** of type **string**
**Description:** Gets and sets the shading color.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableRow-shadingColor)_

**Resource name:** [tableRow](tablerow.md)
**What's new:** Property **values** of type **string**
**Description:** Gets and sets the text values in the row, as a 1D Javascript array.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableRow-values)_

**Resource name:** [tableRow](tablerow.md)
**What's new:** Relationship **cellPaddingBottom** of type **[float](float.md)**
**Description:** Gets and sets the default bottom cell padding for the row in points.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableRow-cellPaddingBottom)_

**Resource name:** [tableRow](tablerow.md)
**What's new:** Relationship **cellPaddingLeft** of type **[float](float.md)**
**Description:** Gets and sets the default left cell padding for the row in points.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableRow-cellPaddingLeft)_

**Resource name:** [tableRow](tablerow.md)
**What's new:** Relationship **cellPaddingRight** of type **[float](float.md)**
**Description:** Gets and sets the default right cell padding for the row in points.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableRow-cellPaddingRight)_

**Resource name:** [tableRow](tablerow.md)
**What's new:** Relationship **cellPaddingTop** of type **[float](float.md)**
**Description:** Gets and sets the default top cell padding for the row in points.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableRow-cellPaddingTop)_

**Resource name:** [tableRow](tablerow.md)
**What's new:** Relationship **cells** of type **[TableCellCollection](tablecellcollection.md)**
**Description:** Gets cells. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableRow-cells)_

**Resource name:** [tableRow](tablerow.md)
**What's new:** Relationship **font** of type **[Font](font.md)**
**Description:** Gets the font. Use this to get and set font name, size, color, and other properties. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableRow-font)_

**Resource name:** [tableRow](tablerow.md)
**What's new:** Relationship **next** of type **[TableRow](tablerow.md)**
**Description:** Gets the next row. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableRow-next)_

**Resource name:** [tableRow](tablerow.md)
**What's new:** Relationship **parentTable** of type **[Table](table.md)**
**Description:** Gets parent table. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableRow-parentTable)_

**Resource name:** [tableRow](tablerow.md)
**What's new:** Relationship **preferredHeight** of type **[float](float.md)**
**Description:** Gets and sets the preferred height of the row in points.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableRow-preferredHeight)_

**Resource name:** [tableRow](tablerow.md)
**What's new:** Relationship **verticalAlignment** of type **[VerticalAlignment](verticalalignment.md)**
**Description:** Gets and sets the vertical alignment of the cells in the row.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableRow-verticalAlignment)_

**Resource name:** [tableRow](tablerow.md)
**What's new:** Method **methodPlusLink** returning **void**
**Description:** Clears the contents of the row.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableRow-clear)_

**Resource name:** [tableRow](tablerow.md)
**What's new:** Method **methodPlusLink** returning **void**
**Description:** Deletes the entire row.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableRow-delete)_

**Resource name:** [tableRow](tablerow.md)
**What's new:** Method **methodPlusLink** returning **[TableBorderStyle](tableborderstyle.md)**
**Description:** Gets the border style of the cells in the row.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableRow-getBorderStyle)_

**Resource name:** [tableRow](tablerow.md)
**What's new:** Method **methodPlusLink** returning **void**
**Description:** Inserts rows using this row as a template. If values are specified, inserts the values into the new rows.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableRow-insertRows)_

**Resource name:** [tableRow](tablerow.md)
**What's new:** Method **methodPlusLink** returning **[TableCell](tablecell.md)**
**Description:** Merges the row into one cell.
**Available in requirement set:** WordApiDesktop, 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableRow-merge)_

**Resource name:** [tableRow](tablerow.md)
**What's new:** Method **methodPlusLink** returning **[SearchResultCollection](searchresultcollection.md)**
**Description:** Performs a search with the specified searchOptions on the scope of the row. The search results are a collection of range objects.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableRow-search)_

**Resource name:** [tableRow](tablerow.md)
**What's new:** Method **methodPlusLink** returning **void**
**Description:** Selects the row and navigates the Word UI to it.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableRow-select)_

**Resource name:** [tableRowCollection](tablerowcollection.md)
**What's new:** Property **items** of type **[TableRow[]](tablerow.md)**
**Description:** A collection of tableRow objects. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableRowCollection-items)_

**Resource name:** [tableRowCollection](tablerowcollection.md)
**What's new:** Relationship **first** of type **[TableRow](tablerow.md)**
**Description:** Gets the first row in this collection. Read-only.
**Available in requirement set:** 1.3
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableRowCollection-first)_

**Resource name:** [tableRowCollection](tablerowcollection.md)
**What's new:** Method **methodPlusLink** returning **[TableRow](tablerow.md)**
**Description:** Gets a table row object by its index in the collection.
**Available in requirement set:** 1.3
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableRowCollection-getItem)_

