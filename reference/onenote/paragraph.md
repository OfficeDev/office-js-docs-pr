# Paragraph Object (JavaScript API for OneNote)

_Applies to: OneNote Online_
_Note: This API is in preview_


Represents a placeholder object for the contents of an outline. A paragraph can contain any one ParagraphType type of content. Paragraphs are automatically positioned.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|id|string|Gets the ID of the paragraph. Read-only.|
|type|string|Gets the paragraph type. Read-only. Possible values are: RichText, Image, Table, InkDrawing, InsertedFile, MediaFile, InkParagraph, Other.|


## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|image|[Image](image.md)|Gets the image in the paragraph. Returns null if ParagraphType is not Image. Read-only.|
|outline|[Outline](outline.md)|Gets the parent outline of the paragraph. Read-only.|
|parentParagraph|[Paragraph](paragraph.md)|Gets the paragraph that contains the paragraph. Returns null if the paragraph is a direct child of the outline. Read-only.|
|richText|[RichText](richtext.md)|Gets the rich text of the paragraph. Returns null if ParagraphType is not RichText. Read-only.|
|subParagraphs|[ParagraphCollection](paragraphcollection.md)|Gets the child paragraphs of the paragraph. Only List and Table paragraph types can have child paragraphs. Read-only.|

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
