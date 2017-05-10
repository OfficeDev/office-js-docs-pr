# ParagraphCollection Object (JavaScript API for Word)

_Word 2016, Word for iPad, Word for Mac, Word Online_

Contains a collection of [paragraph](paragraph.md) objects.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[Paragraph[]](paragraph.md)|A collection of paragraph objects. Read-only.|[1.1](../requirement-sets/word-api-requirement-sets.md)|

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getFirst()](#getfirst)|[Paragraph](paragraph.md)|Gets the first paragraph in this collection. Throws if the collection is empty.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[getFirstOrNullObject()](#getfirstornullobject)|[Paragraph](paragraph.md)|Gets the first paragraph in this collection. Returns a null object if the collection is empty.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[getItem(index: number)](#getitemindex-number)|[Paragraph](paragraph.md)|Gets a paragraph object by its index in the collection.|[1.1](../requirement-sets/word-api-requirement-sets.md)|
|[getLast()](#getlast)|[Paragraph](paragraph.md)|Gets the last paragraph in this collection. Throws if the collection is empty.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[getLastOrNullObject()](#getlastornullobject)|[Paragraph](paragraph.md)|Gets the last paragraph in this collection. Returns a null object if the collection is empty.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[1.1](../requirement-sets/word-api-requirement-sets.md)|

## Method Details


### getFirst()
Gets the first paragraph in this collection. Throws if the collection is empty.

#### Syntax
```js
paragraphCollectionObject.getFirst();
```

#### Parameters
None

#### Returns
[Paragraph](paragraph.md)

### getFirstOrNullObject()
Gets the first paragraph in this collection. Returns a null object if the collection is empty.

#### Syntax
```js
paragraphCollectionObject.getFirstOrNullObject();
```

#### Parameters
None

#### Returns
[Paragraph](paragraph.md)

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

### getLast()
Gets the last paragraph in this collection. Throws if the collection is empty.

#### Syntax
```js
paragraphCollectionObject.getLast();
```

#### Parameters
None

#### Returns
[Paragraph](paragraph.md)

### getLastOrNullObject()
Gets the last paragraph in this collection. Returns a null object if the collection is empty.

#### Syntax
```js
paragraphCollectionObject.getLastOrNullObject();
```

#### Parameters
None

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
