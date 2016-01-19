# InlinePicture object (JavaScript API for Word)

Represents an inline picture.

_Applies to: Word 2016, Word for iPad, Word for Mac_

## Properties
| Property	   | Type	|Description
|:---------------|:--------|:----------|
|altTextDescription|string|Gets or sets a string that represents the alternative text associated with the inline image|
|altTextTitle|string|Gets or sets a string that contains the title for the inline image.|
|hyperlink|string|Gets or sets the hyperlink associated with the inline image.|
|lockAspectRatio|bool|Gets or sets a value that indicates whether the inline image retains its original proportions when you resize it.|


_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|height|**float**|Gets or sets a number that describes the height of the inline image. This is measured in points. |
|parentContentControl|[ContentControl](contentcontrol.md)|Gets the content control that contains the inline image. Returns null if there isn't a parent content control. Read-only.|
|paragraph|[paragraph](paragraph.md)|Gets the paragraph that contains the inline image. Read-only.
|width|**float**|Gets or sets a number that describes the width of the inline image. This is measured in points.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[delete()](#delete)|void|Deletes the picture from the document.|
|[getBase64ImageSrc()](#getbase64imagesrc)|string|Gets an objects whose value is the base64 encoded string representation of the inline image.|
|[insertBreak(breakType: BreakType, insertLocation: InsertLocation)](#insertbreakbreaktype-breaktype-insertlocation-insertlocation)|void|Inserts a break at the specified location. The insertLocation value can be 'Before' or 'After'.|
|[insertContentControl()](#insertcontentcontrol)|[ContentControl](contentcontrol.md)|Wraps the inline picture with a rich text content control.|
|[insertFileFromBase64(base64File: string, insertLocation: InsertLocation)](#insertfilefrombase64base64file-string-insertlocation-insertlocation)|[Range](range.md)|Inserts a document into the body at the specified location. The insertLocation value can be 'Before' or 'After'.|
|[insertHtml(html: string, insertLocation: InsertLocation)](#inserthtmlhtml-string-insertlocation-insertlocation)|[Range](range.md)|Inserts HTML at the specified location. The insertLocation value can be 'Before' or 'After'.|
|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)](#insertInlinePictureFromBase64base64EncodedImage-string-insertlocation-insertlocation)|[InlinePicture](inlinepicture.md)|Inserts a picture into the body at the specified location. The insertLocation value can be 'Replace', 'Before' or 'After'. |
|[insertOoxml(ooxml: string, insertLocation: InsertLocation)](#insertooxmlooxml-string-insertlocation-insertlocation)|[Range](range.md)|Inserts OOXML at the specified location.  The insertLocation value can be 'Before' or 'After'.|
|[insertParagraph(paragraphText: string, insertLocation: InsertLocation)](#insertparagraphparagraphtext-string-insertlocation-insertlocation)|[Paragraph](paragraph.md)|Inserts a paragraph at the specified location. The insertLocation value can be 'Before' or 'After'.|
|[insertText(text: string, insertLocation: InsertLocation)](#inserttexttext-string-insertlocation-insertlocation)|[Range](range.md)|Inserts text into the body at the specified location. The insertLocation value can be 'Before' or 'After'.|
|[select(selectionMode: SelectionMode)](#selectselectionmode-selectionmode)|void|Selects the picture and navigates the Word UI to it. The selectionMode values can be 'Select', 'Start', or 'End'.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method details

### delete()
Deletes the picture from the document.

#### Syntax
```js
inlinePictureObject.delete();
```

#### Parameters
None

#### Returns
void

### getBase64ImageSrc()
Gets an objects whose value is the base64 encoded string representation of the inline image.

#### Syntax
```js
inlinePictureObject.getBase64ImageSrc();
```

#### Parameters
None

#### Returns
string

### insertBreak(breakType: BreakType, insertLocation: InsertLocation)

#### Syntax
```js
inlinePictureObject.insertBreak(breakType, insertLocation);
```
#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|breakType|BreakType|Required. The break type to add to the body.|
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
Inserts a document into the body at the specified location. The insertLocation value can be 'Before' or 'After'.

#### Syntax
```js
inlinePictureObject.insertFileFromBase64(base64File, insertLocation);
```
#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|base64File|string|Required. The base64 encoded contents of a docx file.|
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
|:---------------|:--------|:----------|
|html|string|Required. The HTML to be inserted in the document.|
|insertLocation|InsertLocation|Required. The value can be 'Before' or 'After'.|

#### Returns
[Range](range.md)


### insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)
Inserts a picture into the body at the specified location. The insertLocation value can be 'Before' or 'After'.

#### Syntax
inlinePictureObject.insertInlinePictureFromBase64(image, insertLocation);

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|base64EncodedImage|string|Required. The base64 encoded image to be inserted in the body.|
|insertLocation|InsertLocation|Required. The value can be 'Before' or 'After'.|

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
|:---------------|:--------|:----------|
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
|:---------------|:--------|:----------|
|paragraphText|string|Required. The paragraph text to be inserted.|
|insertLocation|InsertLocation|Required. The value can be 'Before' or 'After'.|

#### Returns
[Paragraph](paragraph.md)

### insertText(text: string, insertLocation: InsertLocation)
Inserts text into the body at the specified location. The insertLocation value can be 'Before' or 'After'.

#### Syntax
```js
inlinePictureObject.insertText(text, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|text|string|Required. Text to be inserted.|
|insertLocation|InsertLocation|Required. The value can be 'Before' or 'After'.|

#### Returns
[Range](range.md)

### select(selectionMode: SelectionMode)
Selects the picture and navigates the Word UI to it. The selectionMode values can be 'Select', 'Start', or 'End'.

#### Syntax
```js
inlinePictureObject.select(selectionMode);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|selectionMode|SelectionMode|Optional. The selection mode can be 'Select', 'Start' or 'End'. 'Select' is the default.|

#### Returns
void

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

## Support details

Use the [requirement set](https://msdn.microsoft.com/EN-US/library/office/mt590206.aspx) in run time checks to make sure your application is supported by the host version of Word. For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](https://msdn.microsoft.com/EN-US/library/office/dn833104.aspx). 