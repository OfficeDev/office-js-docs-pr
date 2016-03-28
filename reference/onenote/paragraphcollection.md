# ParagraphCollection Object (JavaScript API for OneNote)

_Applies to: OneNote Online_
_Note: This API is in public preview_

Represents a collection of paragraphs.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|items|[Paragraph[]](paragraph.md)|A collection of paragraph objects. Read-only.|


## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getItem(index: number or string)](#getitemindex-number-or-string)|[Paragraph](paragraph.md)|Gets a paragraph by its index in the collection. Read-only.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### getItem(index: number or string)
Gets a paragraph by its index in the collection. Read-only.

#### Syntax
```js
paragraphCollectionObject.getItem(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number or string|A number or ID that identifies the index location of the paragraph.|

#### Returns
[Paragraph](paragraph.md)

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
