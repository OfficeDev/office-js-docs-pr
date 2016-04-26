# InlinePicture Object (JavaScript API for Word)

_Word 2016, Word for iPad, Word for Mac_

Represents an inline picture.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|altTextDescription|string|Gets or sets a string that represents the alternative text associated with the inline image|1.1||
|altTextTitle|string|Gets or sets a string that contains the title for the inline image.|1.1||
|hyperlink|string|Gets or sets the hyperlink associated with the inline image.|1.1||
|lockAspectRatio|bool|Gets or sets a value that indicates whether the inline image retains its original proportions when you resize it.|1.1||

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|height|[float](float.md)|Gets or sets a number that describes the height of the inline image.|1.1||
|imageFormat|[ImageFormat](imageformat.md)|Gets the format of the inline image. Read-only.|1.3||
|next|[InlinePicture](inlinepicture.md)|Gets the next inline image. Read-only.|1.3||
|paragraph|[Paragraph](paragraph.md)|Gets the paragraph that contains the inline image. Read-only.|1.2||
|parentContentControl|[ContentControl](contentcontrol.md)|Gets the content control that contains the inline image. Returns null if there isn't a parent content control. Read-only.|1.1||
|parentTable|[Table](table.md)|Gets the table that contains the inline image. Returns null if it is not contained in a table. Read-only.|1.3||
|parentTableCell|[TableCell](tablecell.md)|Gets the table cell that contains the inline image. Returns null if it is not contained in a table cell. Read-only.|1.3||
|width|[float](float.md)|Gets or sets a number that describes the width of the inline image.|1.1||

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|Deletes the inline picture from the document.|1.2|
|[getBase64ImageSrc()](#getbase64imagesrc)|string|Gets the base64 encoded string representation of the inline image.|1.1|
|[getRange(rangeLocation: RangeLocation)](#getrangerangelocation-rangelocation)|[Range](range.md)|Gets the picture, or the starting or ending point of the picture, as a range.|1.3|
|[insertBreak(breakType: BreakType, insertLocation: InsertLocation)](#insertbreakbreaktype-breaktype-insertlocation-insertlocation)|void|Inserts a break at the specified location in the main document. The insertLocation value can be 'Before' or 'After'.|1.2|
|[insertContentControl()](#insertcontentcontrol)|[ContentControl](contentcontrol.md)|Wraps the inline picture with a rich text content control.|1.1|
|[insertFileFromBase64(base64File: string, insertLocation: InsertLocation)](#insertfilefrombase64base64file-string-insertlocation-insertlocation)|[Range](range.md)|Inserts a document at the specified location. The insertLocation value can be 'Before' or 'After'.|1.2|
|[insertHtml(html: string, insertLocation: InsertLocation)](#inserthtmlhtml-string-insertlocation-insertlocation)|[Range](range.md)|Inserts HTML at the specified location. The insertLocation value can be 'Before' or 'After'.|1.2|
|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)](#insertinlinepicturefrombase64base64encodedimage-string-insertlocation-insertlocation)|[InlinePicture](inlinepicture.md)|Inserts an inline picture at the specified location. The insertLocation value can be 'Replace', 'Before' or 'After'.|1.2|
|[insertOoxml(ooxml: string, insertLocation: InsertLocation)](#insertooxmlooxml-string-insertlocation-insertlocation)|[Range](range.md)|Inserts OOXML at the specified location.  The insertLocation value can be 'Before' or 'After'.|1.2|
|[insertParagraph(paragraphText: string, insertLocation: InsertLocation)](#insertparagraphparagraphtext-string-insertlocation-insertlocation)|[Paragraph](paragraph.md)|Inserts a paragraph at the specified location. The insertLocation value can be 'Before' or 'After'.|1.2|
|[insertText(text: string, insertLocation: InsertLocation)](#inserttexttext-string-insertlocation-insertlocation)|[Range](range.md)|Inserts text at the specified location. The insertLocation value can be 'Before' or 'After'.|1.2|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.1|
|[select(selectionMode: SelectionMode)](#selectselectionmode-selectionmode)|void|Selects the inline picture. This causes Word to scroll to the selection.|1.2|

## Method Details


### delete()
Deletes the inline picture from the document.

#### Syntax
```js
inlinePictureObject.delete();
```

#### Parameters
None

#### Returns
void

### getBase64ImageSrc()
Gets the base64 encoded string representation of the inline image.

#### Syntax
```js
inlinePictureObject.getBase64ImageSrc();
```

#### Parameters
None

#### Returns
string

### getRange(rangeLocation: RangeLocation)
Gets the picture, or the starting or ending point of the picture, as a range.

#### Syntax
```js
inlinePictureObject.getRange(rangeLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|rangeLocation|RangeLocation|Optional. Optional. The range location can be 'Whole', 'Start' or 'End'.|

#### Returns
[Range](range.md)

### insertBreak(breakType: BreakType, insertLocation: InsertLocation)
Inserts a break at the specified location in the main document. The insertLocation value can be 'Before' or 'After'.

#### Syntax
```js
inlinePictureObject.insertBreak(breakType, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|breakType|BreakType|Required. The break type to add.|
|insertLocation|InsertLocation|Required. The value can be 'Before' or 'After'.|

#### Returns
void

### insertContentControl()
Wraps the inline picture with a rich text content control.

#### Syntax
```js
inlinePictureObject.insertContentControl();
```

#### Parameters
None

#### Returns
[ContentControl](contentcontrol.md)

### insertFileFromBase64(base64File: string, insertLocation: InsertLocation)
Inserts a document at the specified location. The insertLocation value can be 'Before' or 'After'.

#### Syntax
```js
inlinePictureObject.insertFileFromBase64(base64File, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|base64File|string|Required. The base64 encoded content of a .docx file.|
|insertLocation|InsertLocation|Required. The value can be 'Before' or 'After'.|

#### Returns
[Range](range.md)

### insertHtml(html: string, insertLocation: InsertLocation)
Inserts HTML at the specified location. The insertLocation value can be 'Before' or 'After'.

#### Syntax
```js
inlinePictureObject.insertHtml(html, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|html|string|Required. The HTML to be inserted.|
|insertLocation|InsertLocation|Required. The value can be 'Before' or 'After'.|

#### Returns
[Range](range.md)

### insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)
Inserts an inline picture at the specified location. The insertLocation value can be 'Replace', 'Before' or 'After'.

#### Syntax
```js
inlinePictureObject.insertInlinePictureFromBase64(base64EncodedImage, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|base64EncodedImage|string|Required. The base64 encoded image to be inserted.|
|insertLocation|InsertLocation|Required. The value can be 'Replace', 'Before' or 'After'.|

#### Returns
[InlinePicture](inlinepicture.md)

### insertOoxml(ooxml: string, insertLocation: InsertLocation)
Inserts OOXML at the specified location.  The insertLocation value can be 'Before' or 'After'.

#### Syntax
```js
inlinePictureObject.insertOoxml(ooxml, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|ooxml|string|Required. The OOXML to be inserted.|
|insertLocation|InsertLocation|Required. The value can be 'Before' or 'After'.|

#### Returns
[Range](range.md)

### insertParagraph(paragraphText: string, insertLocation: InsertLocation)
Inserts a paragraph at the specified location. The insertLocation value can be 'Before' or 'After'.

#### Syntax
```js
inlinePictureObject.insertParagraph(paragraphText, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|paragraphText|string|Required. The paragraph text to be inserted.|
|insertLocation|InsertLocation|Required. The value can be 'Before' or 'After'.|

#### Returns
[Paragraph](paragraph.md)

### insertText(text: string, insertLocation: InsertLocation)
Inserts text at the specified location. The insertLocation value can be 'Before' or 'After'.

#### Syntax
```js
inlinePictureObject.insertText(text, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|text|string|Required. Text to be inserted.|
|insertLocation|InsertLocation|Required. The value can be 'Before' or 'After'.|

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

### select(selectionMode: SelectionMode)
Selects the inline picture. This causes Word to scroll to the selection.

#### Syntax
```js
inlinePictureObject.select(selectionMode);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|selectionMode|SelectionMode|Optional. Optional. The selection mode can be 'Select', 'Start' or 'End'. 'Select' is the default.|

#### Returns
void
