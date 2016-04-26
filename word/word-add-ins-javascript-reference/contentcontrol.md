# ContentControl Object (JavaScript API for Word)

_Word 2016, Word for iPad, Word for Mac_

Represents a content control. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as images, tables, or paragraphs of formatted text. Currently, only rich text content controls are supported.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|cannotDelete|bool|Gets or sets a value that indicates whether the user can delete the content control. Mutually exclusive with removeWhenEdited.|WordApi1.1||
|cannotEdit|bool|Gets or sets a value that indicates whether the user can edit the contents of the content control.|WordApi1.1||
|color|string|Gets or sets the color of the content control. Color is specified in '#RRGGBB' format or by using the color name.|WordApi1.1||
|placeholderText|string|Gets or sets the placeholder text of the content control. Dimmed text will be displayed when the content control is empty.|WordApi1.1||
|removeWhenEdited|bool|Gets or sets a value that indicates whether the content control is removed after it is edited. Mutually exclusive with cannotDelete.|WordApi1.1||
|style|string|Gets or sets the style used for the content control. This is the name of the pre-installed or custom style.|WordApi1.1||
|tag|string|Gets or sets a tag to identify a content control.|WordApi1.1||
|text|string|Gets the text of the content control. Read-only.|WordApi1.1||
|title|string|Gets or sets the title for a content control.|WordApi1.1||

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|appearance|[ContentControlAppearance](contentcontrolappearance.md)|Gets or sets the appearance of the content control. The value can be 'boundingBox', 'tags' or 'hidden'.|WordApi1.1||
|contentControls|[ContentControlCollection](contentcontrolcollection.md)|Gets the collection of content control objects in the content control. Read-only.|WordApi1.1||
|font|[Font](font.md)|Gets the text format of the content control. Use this to get and set font name, size, color, and other properties. Read-only.|WordApi1.1||
|id|[uint](uint.md)|Gets an integer that represents the content control identifier. Read-only.|WordApi1.1||
|inlinePictures|[InlinePictureCollection](inlinepicturecollection.md)|Gets the collection of inlinePicture objects in the content control. The collection does not include floating images. Read-only.|WordApi1.1||
|lists|[ListCollection](listcollection.md)|Gets the collection of list objects in the content control. Read-only.|WordApi1.3||
|paragraphs|[ParagraphCollection](paragraphcollection.md)|Get the collection of paragraph objects in the content control. Read-only.|WordApi1.1||
|parentContentControl|[ContentControl](contentcontrol.md)|Gets the content control that contains the content control. Returns null if there isn't a parent content control. Read-only.|WordApi1.1||
|parentTable|[Table](table.md)|Gets the table that contains the content control. Returns null if it is not contained in a table. Read-only.|WordApi1.3||
|parentTableCell|[TableCell](tablecell.md)|Gets the table cell that contains the content control. Returns null if it is not contained in a table cell. Read-only.|WordApi1.3||
|subtype|[ContentControlType](contentcontroltype.md)|Gets the content control subtype. The subtype can be 'RichTextInline', 'RichTextParagraphs', 'RichTextTableCell', 'RichTextTableRow' and 'RichTextTable' for rich text content controls. Read-only.|WordApi1.3||
|tables|[TableCollection](tablecollection.md)|Gets the collection of table objects in the content control. Read-only.|WordApi1.3||
|type|[ContentControlType](contentcontroltype.md)|Gets the content control type. Only rich text content controls are supported currently. Read-only.|WordApi1.1||

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[clear()](#clear)|void|Clears the contents of the content control. The user can perform the undo operation on the cleared content.|WordApi1.1|
|[delete(keepContent: bool)](#deletekeepcontent-bool)|void|Deletes the content control and its content. If keepContent is set to true, the content is not deleted.|WordApi1.1|
|[getHtml()](#gethtml)|string|Gets the HTML representation of the content control object.|WordApi1.1|
|[getOoxml()](#getooxml)|string|Gets the Office Open XML (OOXML) representation of the content control object.|WordApi1.1|
|[getRange(rangeLocation: RangeLocation)](#getrangerangelocation-rangelocation)|[Range](range.md)|Gets the whole content control, or the starting or ending point of the content control, as a range.|WordApi1.3|
|[getTextRanges(punctuationMarks: string[], trimSpacing: bool)](#gettextrangespunctuationmarks-string-trimspacing-bool)|[RangeCollection](rangecollection.md)|Gets the text ranges in the content control by using punctuation marks andor space character.|WordApi1.3|
|[insertBreak(breakType: BreakType, insertLocation: InsertLocation)](#insertbreakbreaktype-breaktype-insertlocation-insertlocation)|void|Inserts a break at the specified location in the main document. The insertLocation value can be 'Start', 'End', 'Before' or 'After'. This method cannot be used with 'RichTextTable', 'RichTextTableRow' and 'RichTextTableCell' content controls.|WordApi1.1|
|[insertFileFromBase64(base64File: string, insertLocation: InsertLocation)](#insertfilefrombase64base64file-string-insertlocation-insertlocation)|[Range](range.md)|Inserts a document into the content control at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.|WordApi1.1|
|[insertHtml(html: string, insertLocation: InsertLocation)](#inserthtmlhtml-string-insertlocation-insertlocation)|[Range](range.md)|Inserts HTML into the content control at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.|WordApi1.1|
|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)](#insertinlinepicturefrombase64base64encodedimage-string-insertlocation-insertlocation)|[InlinePicture](inlinepicture.md)|Inserts an inline picture into the content control at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.|WordApi1.2|
|[insertOoxml(ooxml: string, insertLocation: InsertLocation)](#insertooxmlooxml-string-insertlocation-insertlocation)|[Range](range.md)|Inserts OOXML into the content control at the specified location.  The insertLocation value can be 'Replace', 'Start' or 'End'.|WordApi1.1|
|[insertParagraph(paragraphText: string, insertLocation: InsertLocation)](#insertparagraphparagraphtext-string-insertlocation-insertlocation)|[Paragraph](paragraph.md)|Inserts a paragraph at the specified location. The insertLocation value can be 'Start', 'End', 'Before' or 'After'.|WordApi1.1|
|[insertTable(rowCount: number, columnCount: number, insertLocation: InsertLocation, values: string[][])](#inserttablerowcount-number-columncount-number-insertlocation-insertlocation-values-string)|[Table](table.md)|Inserts a table with the specified number of rows and columns into, or next to, a content control. The insertLocation value can be 'Start', 'End', 'Before' or 'After'.|WordApi1.3|
|[insertText(text: string, insertLocation: InsertLocation)](#inserttexttext-string-insertlocation-insertlocation)|[Range](range.md)|Inserts text into the content control at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.|WordApi1.1|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|WordApi1.1|
|[search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)](#searchsearchtext-string-searchoptions-paramtypestrings.searchoptions)|[SearchResultCollection](searchresultcollection.md)|Performs a search with the specified searchOptions on the scope of the content control object. The search results are a collection of range objects.|WordApi1.1|
|[select(selectionMode: SelectionMode)](#selectselectionmode-selectionmode)|void|Selects the content control. This causes Word to scroll to the selection.|WordApi1.1|
|[split(delimiters: string[], multiParagraphs: bool, trimDelimiters: bool, trimSpacing: bool)](#splitdelimiters-string-multiparagraphs-bool-trimdelimiters-bool-trimspacing-bool)|[RangeCollection](rangecollection.md)|Splits the content control into child ranges by using delimiters.|WordApi1.3|

## Method Details


### clear()
Clears the contents of the content control. The user can perform the undo operation on the cleared content.

#### Syntax
```js
contentControlObject.clear();
```

#### Parameters
None

#### Returns
void

### delete(keepContent: bool)
Deletes the content control and its content. If keepContent is set to true, the content is not deleted.

#### Syntax
```js
contentControlObject.delete(keepContent);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|keepContent|bool|Required. Indicates whether the content should be deleted with the content control. If keepContent is set to true, the content is not deleted.|

#### Returns
void

### getHtml()
Gets the HTML representation of the content control object.

#### Syntax
```js
contentControlObject.getHtml();
```

#### Parameters
None

#### Returns
string

### getOoxml()
Gets the Office Open XML (OOXML) representation of the content control object.

#### Syntax
```js
contentControlObject.getOoxml();
```

#### Parameters
None

#### Returns
string

### getRange(rangeLocation: RangeLocation)
Gets the whole content control, or the starting or ending point of the content control, as a range.

#### Syntax
```js
contentControlObject.getRange(rangeLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|rangeLocation|RangeLocation|Optional. Optional. The range location can be 'Whole', 'Start' or 'End'.|

#### Returns
[Range](range.md)

### getTextRanges(punctuationMarks: string[], trimSpacing: bool)
Gets the text ranges in the content control by using punctuation marks andor space character.

#### Syntax
```js
contentControlObject.getTextRanges(punctuationMarks, trimSpacing);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|punctuationMarks|string[]|Required. The punctuation marks and/or space character as an array of strings.|
|trimSpacing|bool|Optional. Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.|

#### Returns
[RangeCollection](rangecollection.md)

### insertBreak(breakType: BreakType, insertLocation: InsertLocation)
Inserts a break at the specified location in the main document. The insertLocation value can be 'Start', 'End', 'Before' or 'After'. This method cannot be used with 'RichTextTable', 'RichTextTableRow' and 'RichTextTableCell' content controls.

#### Syntax
```js
contentControlObject.insertBreak(breakType, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|breakType|BreakType|Required. Type of break.|
|insertLocation|InsertLocation|Required. The value can be 'Start', 'End', 'Before' or 'After'.|

#### Returns
void

### insertFileFromBase64(base64File: string, insertLocation: InsertLocation)
Inserts a document into the content control at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.

#### Syntax
```js
contentControlObject.insertFileFromBase64(base64File, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|base64File|string|Required. The base64 encoded content of a .docx file.|
|insertLocation|InsertLocation|Required. The value can be 'Replace', 'Start' or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.|

#### Returns
[Range](range.md)

### insertHtml(html: string, insertLocation: InsertLocation)
Inserts HTML into the content control at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.

#### Syntax
```js
contentControlObject.insertHtml(html, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|html|string|Required. The HTML to be inserted in to the content control.|
|insertLocation|InsertLocation|Required. The value can be 'Replace', 'Start' or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.|

#### Returns
[Range](range.md)

### insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)
Inserts an inline picture into the content control at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.

#### Syntax
```js
contentControlObject.insertInlinePictureFromBase64(base64EncodedImage, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|base64EncodedImage|string|Required. The base64 encoded image to be inserted in the content control.|
|insertLocation|InsertLocation|Required. The value can be 'Replace', 'Start' or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.|

#### Returns
[InlinePicture](inlinepicture.md)

### insertOoxml(ooxml: string, insertLocation: InsertLocation)
Inserts OOXML into the content control at the specified location.  The insertLocation value can be 'Replace', 'Start' or 'End'.

#### Syntax
```js
contentControlObject.insertOoxml(ooxml, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|ooxml|string|Required. The OOXML to be inserted in to the content control.|
|insertLocation|InsertLocation|Required. The value can be 'Replace', 'Start' or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.|

#### Returns
[Range](range.md)

### insertParagraph(paragraphText: string, insertLocation: InsertLocation)
Inserts a paragraph at the specified location. The insertLocation value can be 'Start', 'End', 'Before' or 'After'.

#### Syntax
```js
contentControlObject.insertParagraph(paragraphText, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|paragraphText|string|Required. The paragrph text to be inserted.|
|insertLocation|InsertLocation|Required. The value can be 'Start', 'End', 'Before' or 'After'. 'Before' and 'After' cannot be used with 'RichTextTable', 'RichTextTableRow' and 'RichTextTableCell' content controls.|

#### Returns
[Paragraph](paragraph.md)

### insertTable(rowCount: number, columnCount: number, insertLocation: InsertLocation, values: string[][])
Inserts a table with the specified number of rows and columns into, or next to, a content control. The insertLocation value can be 'Start', 'End', 'Before' or 'After'.

#### Syntax
```js
contentControlObject.insertTable(rowCount, columnCount, insertLocation, values);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|rowCount|number|Required. The number of rows in the table.|
|columnCount|number|Required. The number of columns in the table.|
|insertLocation|InsertLocation|Required. The value can be 'Start', 'End', 'Before' or 'After'. 'Before' and 'After' cannot be used with 'RichTextTable', 'RichTextTableRow' and 'RichTextTableCell' content controls.|
|values|string[][]|Optional. Optional 2D array. Cells are filled if the corresponding strings are specified in the array.|

#### Returns
[Table](table.md)

### insertText(text: string, insertLocation: InsertLocation)
Inserts text into the content control at the specified location. The insertLocation value can be 'Replace', 'Start' or 'End'.

#### Syntax
```js
contentControlObject.insertText(text, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|text|string|Required. The text to be inserted in to the content control.|
|insertLocation|InsertLocation|Required. The value can be 'Replace', 'Start' or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.|

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
Performs a search with the specified searchOptions on the scope of the content control object. The search results are a collection of range objects.

#### Syntax
```js
contentControlObject.search(searchText, searchOptions);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|searchText|string|Required. The search text.|
|searchOptions|ParamTypeStrings.SearchOptions|Optional. Optional. Options for the search.|

#### Returns
[SearchResultCollection](searchresultcollection.md)

### select(selectionMode: SelectionMode)
Selects the content control. This causes Word to scroll to the selection.

#### Syntax
```js
contentControlObject.select(selectionMode);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|selectionMode|SelectionMode|Optional. Optional. The selection mode can be 'Select', 'Start' or 'End'. 'Select' is the default.|

#### Returns
void

### split(delimiters: string[], multiParagraphs: bool, trimDelimiters: bool, trimSpacing: bool)
Splits the content control into child ranges by using delimiters.

#### Syntax
```js
contentControlObject.split(delimiters, multiParagraphs, trimDelimiters, trimSpacing);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|delimiters|string[]|Required. The delimiters as an array of strings.|
|multiParagraphs|bool|Optional. Optional. Indicates whether a returned child range can cover multiple paragraphs. Default is false which indicates that the paragraph boundaries are also used as delimiters.|
|trimDelimiters|bool|Optional. Optional. Indicates whether to trim delimiters from the ranges in the range collection. Default is false which indicates that the delimiters are included in the ranges returned in the range collection.|
|trimSpacing|bool|Optional. Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.|

#### Returns
[RangeCollection](rangecollection.md)
