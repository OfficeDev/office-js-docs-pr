# Paragraph Object (JavaScript API for OneNote)

_Applies to: OneNote Online_
_Note: This API is in preview_

A container for the visible content on a page. A Paragraph can contain any one ParagraphType type of content, such as RichText, Image, or Table.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|id|string|Gets the ID of the Paragraph object. Read-only.|
|type|string|Gets the type of the Paragraph object. Read-only. Possible values are: RichText, Image, Table, InkDrawing, InsertedFile, MediaFile, InkParagraph, Other.|



## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|image|[Image](image.md)|Gets the Image object in the Paragraph. Returns null if ParagraphType is not Image. Read-only.|
|outline|[Outline](outline.md)|Gets the Outline object that contains the Paragraph. Read-only.|
|parentParagraph|[Paragraph](paragraph.md)|Gets the Paragraph object that contains the Paragraph. Returns null if the Paragraph is a direct child of an Outline. Read-only.|
|richText|[RichText](richtext.md)|Gets the RichText object in the Paragraph. Returns null if ParagraphType is not RichText. Read-only Read-only.|
|subParagraphs|[ParagraphCollection](paragraphcollection.md)|Gets the child Paragraph objects of the Paragraph. Applies only if ParagraphType is Table. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


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
