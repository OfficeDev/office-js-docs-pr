# InlinePicture Object (JavaScript API for Word)

_Word 2016, Word for iPad, Word for Mac_

Represents an inline picture.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|altTextDescription|string|Gets or sets a string that represents the alternative text associated with the inline image|
|altTextTitle|string|Gets or sets a string that contains the title for the inline image.|
|hyperlink|string|Gets or sets the hyperlink associated with the inline image.|
|imageFormat|string|Gets the format of the inline image. Read-only.|
|lockAspectRatio|bool|Gets or sets a value that indicates whether the inline image retains its original proportions when you resize it.|

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|height|[float](float.md)|Gets or sets a number that describes the height of the inline image.|
|next|[InlinePicture](inlinepicture.md)|Gets the next inline image. Read-only.|
|paragraph|[Paragraph](paragraph.md)|Gets the paragraph that contains the inline image. Read-only.|
|parentContentControl|[ContentControl](contentcontrol.md)|Gets the content control that contains the inline image. Returns null if there isn't a parent content control. Read-only.|
|parentTable|[Table](table.md)|Gets the table that contains the inline image. Returns null if it is not contained in a table. Read-only.|
|parentTableCell|[TableCell](tablecell.md)|Gets the table cell that contains the inline image. Returns null if it is not contained in a table cell. Read-only.|
|width|[float](float.md)|Gets or sets a number that describes the width of the inline image.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[delete()](#delete)|void|Deletes the inline picture from the document.|
|[getBase64ImageSrc()](#getbase64imagesrc)|string|Gets the base64 encoded string representation of the inline image.|
|[getRange(rangeLocation: string)](#getrangerangelocation-string)|[Range](range.md)|Gets the picture, or the starting or ending point of the picture, as a range.|
|[insertBreak(breakType: string, insertLocation: string)](#insertbreakbreaktype-string-insertlocation-string)|void|Inserts a break at the specified location in the main document. The insertLocation value can be 'Before' or 'After'.|
|[insertContentControl()](#insertcontentcontrol)|[ContentControl](contentcontrol.md)|Wraps the inline picture with a rich text content control.|
|[insertFileFromBase64(base64File: string, insertLocation: string)](#insertfilefrombase64base64file-string-insertlocation-string)|[Range](range.md)|Inserts a document at the specified location. The insertLocation value can be 'Before' or 'After'.|
|[insertHtml(html: string, insertLocation: string)](#inserthtmlhtml-string-insertlocation-string)|[Range](range.md)|Inserts HTML at the specified location. The insertLocation value can be 'Before' or 'After'.|
|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: string)](#insertinlinepicturefrombase64base64encodedimage-string-insertlocation-string)|[InlinePicture](inlinepicture.md)|Inserts an inline picture at the specified location. The insertLocation value can be 'Replace', 'Before' or 'After'.|
|[insertOoxml(ooxml: string, insertLocation: string)](#insertooxmlooxml-string-insertlocation-string)|[Range](range.md)|Inserts OOXML at the specified location.  The insertLocation value can be 'Before' or 'After'.|
|[insertParagraph(paragraphText: string, insertLocation: string)](#insertparagraphparagraphtext-string-insertlocation-string)|[Paragraph](paragraph.md)|Inserts a paragraph at the specified location. The insertLocation value can be 'Before' or 'After'.|
|[insertText(text: string, insertLocation: string)](#inserttexttext-string-insertlocation-string)|[Range](range.md)|Inserts text at the specified location. The insertLocation value can be 'Before' or 'After'.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|
|[select(selectionMode: string)](#selectselectionmode-string)|void|Selects the inline picture. This causes Word to scroll to the selection.|

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

### getRange(rangeLocation: string)
Gets the picture, or the starting or ending point of the picture, as a range.

#### Syntax
```js
inlinePictureObject.getRange(rangeLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|rangeLocation|string|Optional. Optional. The range location can be 'Whole', 'Start' or 'End'.  Possible values are: Whole, Start, End|

#### Returns
[Range](range.md)

### insertBreak(breakType: string, insertLocation: string)
Inserts a break at the specified location in the main document. The insertLocation value can be 'Before' or 'After'.

#### Syntax
```js
inlinePictureObject.insertBreak(breakType, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|breakType|string|Required. The break type to add. Possible values are: `Page` Page break at the insertion point.,`Column` Column break at the insertion point.,`Next` Section break on next page.,`SectionContinuous` New section without a corresponding page break.,`SectionEven` Section break with the next section beginning on the next even-numbered page. If the section break falls on an even-numbered page, Word leaves the next odd-numbered page blank.,`SectionOdd` Section break with the next section beginning on the next odd-numbered page. If the section break falls on an odd-numbered page, Word leaves the next even-numbered page blank.,`Line` Line break.,`LineClearLeft` Line break.,`LineClearRight` Line break.,`TextWrapping` Ends the current line and forces the text to continue below a picture, table, or other item. The text continues on the next blank line that does not contain a table aligned with the left or right margin.|
|insertLocation|string|Required. The value can be 'Before' or 'After'. Possible values are: `Before` Add content before the contents of the calling object.,`After` Add content after the contents of the calling object.,`Start` Prepend content to the contents of the calling object.,`End` Append content to the contents of the calling object.,`Replace` Replace the contents of the current object.|

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

### insertFileFromBase64(base64File: string, insertLocation: string)
Inserts a document at the specified location. The insertLocation value can be 'Before' or 'After'.

#### Syntax
```js
inlinePictureObject.insertFileFromBase64(base64File, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|base64File|string|Required. The base64 encoded content of a .docx file.|
|insertLocation|string|Required. The value can be 'Before' or 'After'. Possible values are: `Before` Add content before the contents of the calling object.,`After` Add content after the contents of the calling object.,`Start` Prepend content to the contents of the calling object.,`End` Append content to the contents of the calling object.,`Replace` Replace the contents of the current object.|

#### Returns
[Range](range.md)

### insertHtml(html: string, insertLocation: string)
Inserts HTML at the specified location. The insertLocation value can be 'Before' or 'After'.

#### Syntax
```js
inlinePictureObject.insertHtml(html, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|html|string|Required. The HTML to be inserted.|
|insertLocation|string|Required. The value can be 'Before' or 'After'. Possible values are: `Before` Add content before the contents of the calling object.,`After` Add content after the contents of the calling object.,`Start` Prepend content to the contents of the calling object.,`End` Append content to the contents of the calling object.,`Replace` Replace the contents of the current object.|

#### Returns
[Range](range.md)

### insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: string)
Inserts an inline picture at the specified location. The insertLocation value can be 'Replace', 'Before' or 'After'.

#### Syntax
```js
inlinePictureObject.insertInlinePictureFromBase64(base64EncodedImage, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|base64EncodedImage|string|Required. The base64 encoded image to be inserted.|
|insertLocation|string|Required. The value can be 'Replace', 'Before' or 'After'. Possible values are: `Before` Add content before the contents of the calling object.,`After` Add content after the contents of the calling object.,`Start` Prepend content to the contents of the calling object.,`End` Append content to the contents of the calling object.,`Replace` Replace the contents of the current object.|

#### Returns
[InlinePicture](inlinepicture.md)

### insertOoxml(ooxml: string, insertLocation: string)
Inserts OOXML at the specified location.  The insertLocation value can be 'Before' or 'After'.

#### Syntax
```js
inlinePictureObject.insertOoxml(ooxml, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|ooxml|string|Required. The OOXML to be inserted.|
|insertLocation|string|Required. The value can be 'Before' or 'After'. Possible values are: `Before` Add content before the contents of the calling object.,`After` Add content after the contents of the calling object.,`Start` Prepend content to the contents of the calling object.,`End` Append content to the contents of the calling object.,`Replace` Replace the contents of the current object.|

#### Returns
[Range](range.md)

### insertParagraph(paragraphText: string, insertLocation: string)
Inserts a paragraph at the specified location. The insertLocation value can be 'Before' or 'After'.

#### Syntax
```js
inlinePictureObject.insertParagraph(paragraphText, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|paragraphText|string|Required. The paragraph text to be inserted.|
|insertLocation|string|Required. The value can be 'Before' or 'After'. Possible values are: `Before` Add content before the contents of the calling object.,`After` Add content after the contents of the calling object.,`Start` Prepend content to the contents of the calling object.,`End` Append content to the contents of the calling object.,`Replace` Replace the contents of the current object.|

#### Returns
[Paragraph](paragraph.md)

### insertText(text: string, insertLocation: string)
Inserts text at the specified location. The insertLocation value can be 'Before' or 'After'.

#### Syntax
```js
inlinePictureObject.insertText(text, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|text|string|Required. Text to be inserted.|
|insertLocation|string|Required. The value can be 'Before' or 'After'. Possible values are: `Before` Add content before the contents of the calling object.,`After` Add content after the contents of the calling object.,`Start` Prepend content to the contents of the calling object.,`End` Append content to the contents of the calling object.,`Replace` Replace the contents of the current object.|

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
|:---------------|:--------|:----------|
|param|object|Optional. Accepts parameter and relationship names as delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns
void

### select(selectionMode: string)
Selects the inline picture. This causes Word to scroll to the selection.

#### Syntax
```js
inlinePictureObject.select(selectionMode);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|selectionMode|string|Optional. Optional. The selection mode can be 'Select', 'Start' or 'End'. 'Select' is the default.  Possible values are: Select, Start, End|

#### Returns
void
