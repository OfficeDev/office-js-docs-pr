# ParagraphCollection Object (JavaScript API for Word)

_Word 2016, Word for iPad, Word for Mac_

Contains a collection of [paragraph](paragraph.md) objects.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|items|[Paragraph[]](paragraph.md)|A collection of paragraph objects. Read-only.|

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|first|[Paragraph](paragraph.md)|Gets the first paragraph in this collection. Read-only.|
|last|[Paragraph](paragraph.md)|Gets the last paragraph in this collection. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getItem(index: number)](#getitemindex-number)|[Paragraph](paragraph.md)|Gets a paragraph object by its index in the collection.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### getItem(index: number)
Gets a paragraph object by its index in the collection.

#### Syntax
```js
paragraphCollectionObject.getItem(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|A number that identifies the index location of a paragraph object.|

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
