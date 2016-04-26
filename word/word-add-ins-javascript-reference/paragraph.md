# Paragraph Object (JavaScript API for Word)

_Word 2016, Word for iPad, Word for Mac_

Represents a single paragraph in a selection, range, content control, or document body.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|listLevel|int|Gets or sets the list level of the paragraph.|1.3||
|outlineLevel|int|Gets or sets the outline level for the paragraph.|WordApiDesktop, 1.3||
|style|string|Gets or sets the style used for the paragraph. This is the name of the pre-installed or custom style.|1.1||
|tableNestingLevel|int|Gets the level of the paragraph's table. It returns 0 if the paragraph is not in a table. Read-only.|1.3||
|text|string|Gets the text of the paragraph. Read-only.|1.1||

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|alignment|[Alignment](alignment.md)|Gets or sets the alignment for a paragraph. The value can  be 'left', 'centered', 'right', or 'justified'.|1.1||
|contentControls|[ContentControlCollection](contentcontrolcollection.md)|Gets the collection of content control objects in the paragraph. Read-only.|1.1||
|firstLineIndent|[float](float.md)|Gets or sets the value, in points, for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.|1.1||
|font|[Font](font.md)|Gets the text format of the paragraph. Use this to get and set font name, size, color, and other properties. Read-only.|1.1||
|inlinePictures|[InlinePictureCollection](inlinepicturecollection.md)|Gets the collection of inlinePicture objects in the paragraph. The collection does not include floating images. Read-only.|1.1||
|leftIndent|[float](float.md)|Gets or sets the left indent value, in points, for the paragraph.|1.1||
|lineSpacing|[float](float.md)|Gets or sets the line spacing, in points, for the specified paragraph. In the Word UI, this value is divided by 12.|1.1||
|lineUnitAfter|[float](float.md)|Gets or sets the amount of spacing, in grid lines. after the paragraph.|1.1||
|lineUnitBefore|[float](float.md)|Gets or sets the amount of spacing, in grid lines, before the paragraph.|1.1||
|list|[List](list.md)|Gets the List to which this paragraph belongs. Returns null if the paragraph is not in a list. Read-only.|1.3||
|next|[Paragraph](paragraph.md)|Gets the next paragraph. Read-only.|1.3||
|parentBody|[Body](body.md)|Gets the parent body of the paragraph. Read-only.|1.3||
|parentContentControl|[ContentControl](contentcontrol.md)|Gets the content control that contains the paragraph. Returns null if there isn't a parent content control. Read-only.|1.1||
|parentTable|[Table](table.md)|Gets the table that contains the paragraph. Returns null if it is not contained in a table. Read-only.|1.3||
|parentTableCell|[TableCell](tablecell.md)|Gets the table cell that contains the paragraph. Returns null if it is not contained in a table cell. Read-only.|1.3||
|previous|[Paragraph](paragraph.md)|Gets the previous paragraph. Read-only.|1.3||
|rightIndent|[float](float.md)|Gets or sets the right indent value, in points, for the paragraph.|1.1||
|spaceAfter|[float](float.md)|Gets or sets the spacing, in points, after the paragraph.|1.1||
|spaceBefore|[float](float.md)|Gets or sets the spacing, in points, before the paragraph.|1.1||

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[clear()](#clear)|void|Clears the contents of the paragraph object. The user can perform the undo operation on the cleared content.|1.1|
|[delete()](#delete)|void|Deletes the paragraph and its content from the document.|1.1|
|[getHtml()](#gethtml)|string|Gets the HTML representation of the paragraph object.|1.1|
|[getOoxml()](#getooxml)|string|Gets the Office Open XML (OOXML) representation of the paragraph object.|1.1|
|[getRange(rangeLocation: RangeLocation)](#getrangerangelocation-rangelocation)|[Range](range.md)|Gets the whole paragraph, or the starting or ending point of the paragraph, as a range.|1.3|
|[getTextRanges(punctuationMarks: string[], trimSpacing: bool)](#gettextrangespunctuationmarks-string-trimspacing-bool)|[RangeCollection](rangecollection.md)|Gets the text ranges in the paragraph by using punctuation marks andor space character.|1.3|
|[insertBreak(breakType: BreakType, insertLocation: InsertLocation)](#insertbreakbreaktype-breaktype-insertlocation-insertlocation)|void|Inserts a break at the specified location in the main document. The insertLocation value can be 'Before' or 'After'.|1.1|
|[insertContentControl()](#insertcontentcontrol)|[ContentControl](contentcontrol.md)|Wraps the paragraph object with a rich text content control.|1.1|
|[insertFileFromBase64(base64File: string, insertLocation: InsertLocation)](#insertfilefrombase64base64file-string-insertlocation-insertlocation)|[Range](range.md)|Inserts a document into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.|1.1|
|[insertHtml(html: string, insertLocation: InsertLocation)](#inserthtmlhtml-string-insertlocation-insertlocation)|[Range](range.md)|Inserts HTML into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.|1.1|
|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)](#insertinlinepicturefrombase64base64encodedimage-string-insertlocation-insertlocation)|[InlinePicture](inlinepicture.md)|Inserts a picture into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.|1.1|
|[insertOoxml(ooxml: string, insertLocation: InsertLocation)](#insertooxmlooxml-string-insertlocation-insertlocation)|[Range](range.md)|Inserts OOXML into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.|1.1|
|[insertParagraph(paragraphText: string, insertLocation: InsertLocation)](#insertparagraphparagraphtext-string-insertlocation-insertlocation)|[Paragraph](paragraph.md)|Inserts a paragraph at the specified location. The insertLocation value can be 'Before' or 'After'.|1.1|
|[insertTable(rowCount: number, columnCount: number, insertLocation: InsertLocation, values: string[][])](#inserttablerowcount-number-columncount-number-insertlocation-insertlocation-values-string)|[Table](table.md)|Inserts a table with the specified number of rows and columns. The insertLocation value can be 'Before' or 'After'.|1.3|
|[insertText(text: string, insertLocation: InsertLocation)](#inserttexttext-string-insertlocation-insertlocation)|[Range](range.md)|Inserts text into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.|1.1|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.1|
|[search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)](#searchsearchtext-string-searchoptions-paramtypestrings.searchoptions)|[SearchResultCollection](searchresultcollection.md)|Performs a search with the specified searchOptions on the scope of the paragraph object. The search results are a collection of range objects.|1.1|
|[select(selectionMode: SelectionMode)](#selectselectionmode-selectionmode)|void|Selects and navigates the Word UI to the paragraph.|1.1|
|[split(delimiters: string[], trimDelimiters: bool, trimSpacing: bool)](#splitdelimiters-string-trimdelimiters-bool-trimspacing-bool)|[RangeCollection](rangecollection.md)|Splits the paragraph into child ranges by using delimiters.|1.3|

## Method Details


### clear()
Clears the contents of the paragraph object. The user can perform the undo operation on the cleared content.

#### Syntax
```js
paragraphObject.clear();
```

#### Parameters
None

#### Returns
void

### delete()
Deletes the paragraph and its content from the document.

#### Syntax
```js
paragraphObject.delete();
```

#### Parameters
None

#### Returns
void

### getHtml()
Gets the HTML representation of the paragraph object.

#### Syntax
```js
paragraphObject.getHtml();
```

#### Parameters
None

#### Returns
string

### getOoxml()
Gets the Office Open XML (OOXML) representation of the paragraph object.

#### Syntax
```js
paragraphObject.getOoxml();
```

#### Parameters
None

#### Returns
string

### getRange(rangeLocation: RangeLocation)
Gets the whole paragraph, or the starting or ending point of the paragraph, as a range.

#### Syntax
```js
paragraphObject.getRange(rangeLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|rangeLocation|RangeLocation|Optional. Optional. The range location can be 'Whole', 'Start' or 'End'.|

#### Returns
[Range](range.md)

### getTextRanges(punctuationMarks: string[], trimSpacing: bool)
Gets the text ranges in the paragraph by using punctuation marks andor space character.

#### Syntax
```js
paragraphObject.getTextRanges(punctuationMarks, trimSpacing);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|punctuationMarks|string[]|Required. The punctuation marks and/or space character as an array of strings.|
|trimSpacing|bool|Optional. Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.|

#### Returns
[RangeCollection](rangecollection.md)

### insertBreak(breakType: BreakType, insertLocation: InsertLocation)
Inserts a break at the specified location in the main document. The insertLocation value can be 'Before' or 'After'.

#### Syntax
```js
paragraphObject.insertBreak(breakType, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|breakType|BreakType|Required. The break type to add to the document.|
|insertLocation|InsertLocation|Required. The value can be 'Before' or 'After'.|

#### Returns
void

### insertContentControl()
Wraps the paragraph object with a rich text content control.

#### Syntax
```js
paragraphObject.insertContentControl();
```

#### Parameters
None

#### Returns
[ContentControl](contentcontrol.md)

### insertFileFromBase64(base64File: string, insertLocation: InsertLocation)
Inserts a document into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.

#### Syntax
```js
paragraphObject.insertFileFromBase64(base64File, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|base64File|string|Required. The base64 encoded content of a .docx file.|
|insertLocation|InsertLocation|Required. The value can be 'Replace', 'Start' or 'End'.|

#### Returns
[Range](range.md)

### insertHtml(html: string, insertLocation: InsertLocation)
Inserts HTML into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.

#### Syntax
```js
paragraphObject.insertHtml(html, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|html|string|Required. The HTML to be inserted in the paragraph.|
|insertLocation|InsertLocation|Required. The value can be 'Replace', 'Start' or 'End'.|

#### Returns
[Range](range.md)

### insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)
Inserts a picture into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.

#### Syntax
```js
paragraphObject.insertInlinePictureFromBase64(base64EncodedImage, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|base64EncodedImage|string|Required. The base64 encoded image to be inserted.|
|insertLocation|InsertLocation|Required. The value can be 'Replace', 'Start' or 'End'.|

#### Returns
[InlinePicture](inlinepicture.md)

### insertOoxml(ooxml: string, insertLocation: InsertLocation)
Inserts OOXML into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.

#### Syntax
```js
paragraphObject.insertOoxml(ooxml, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|ooxml|string|Required. The OOXML to be inserted in the paragraph.|
|insertLocation|InsertLocation|Required. The value can be 'Replace', 'Start' or 'End'.|

#### Returns
[Range](range.md)

### insertParagraph(paragraphText: string, insertLocation: InsertLocation)
Inserts a paragraph at the specified location. The insertLocation value can be 'Before' or 'After'.

#### Syntax
```js
paragraphObject.insertParagraph(paragraphText, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|paragraphText|string|Required. The paragraph text to be inserted.|
|insertLocation|InsertLocation|Required. The value can be 'Before' or 'After'.|

#### Returns
[Paragraph](paragraph.md)

### insertTable(rowCount: number, columnCount: number, insertLocation: InsertLocation, values: string[][])
Inserts a table with the specified number of rows and columns. The insertLocation value can be 'Before' or 'After'.

#### Syntax
```js
paragraphObject.insertTable(rowCount, columnCount, insertLocation, values);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|rowCount|number|Required. The number of rows in the table.|
|columnCount|number|Required. The number of columns in the table.|
|insertLocation|InsertLocation|Required. The value can be 'Before' or 'After'.|
|values|string[][]|Optional. Optional 2D array. Cells are filled if the corresponding strings are specified in the array.|

#### Returns
[Table](table.md)

### insertText(text: string, insertLocation: InsertLocation)
Inserts text into the paragraph at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.

#### Syntax
```js
paragraphObject.insertText(text, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|text|string|Required. Text to be inserted.|
|insertLocation|InsertLocation|Required. The value can be 'Replace', 'Start' or 'End'.|

#### Returns
[Range](range.md)

### load(param: object)
Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.

#### Syntax
```js
object.load(param);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|param|object|Optional. Accepts parameter and relationship names as delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns
void

### search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)
Performs a search with the specified searchOptions on the scope of the paragraph object. The search results are a collection of range objects.

#### Syntax
```js
paragraphObject.search(searchText, searchOptions);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|searchText|string|Required. The search text.|
|searchOptions|ParamTypeStrings.SearchOptions|Optional. Optional. Options for the search.|

#### Returns
[SearchResultCollection](searchresultcollection.md)

### select(selectionMode: SelectionMode)
Selects and navigates the Word UI to the paragraph.

#### Syntax
```js
paragraphObject.select(selectionMode);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|selectionMode|SelectionMode|Optional. Optional. The selection mode can be 'Select', 'Start' or 'End'. 'Select' is the default.|

#### Returns
void

### split(delimiters: string[], trimDelimiters: bool, trimSpacing: bool)
Splits the paragraph into child ranges by using delimiters.

#### Syntax
```js
paragraphObject.split(delimiters, trimDelimiters, trimSpacing);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|delimiters|string[]|Required. The delimiters as an array of strings.|
|trimDelimiters|bool|Optional. Optional. Indicates whether to trim delimiters from the ranges in the range collection. Default is false which indicates that the delimiters are included in the ranges returned in the range collection.|
|trimSpacing|bool|Optional. Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.|

#### Returns
[RangeCollection](rangecollection.md)
