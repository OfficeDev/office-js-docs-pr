# InlinePictureCollection Object (JavaScript API for Word)

_Word 2016, Word for iPad, Word for Mac_

Contains a collection of [inlinePicture](inlinePicture.md) objects.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|items|[InlinePicture[]](inlinepicture.md)|A collection of inlinePicture objects. Read-only.|

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|first|[InlinePicture](inlinepicture.md)|Gets the first inline image in this collection. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getItem(index: number)](#getitemindex-number)|[InlinePicture](inlinepicture.md)|Gets an inline picture object by its index in the collection.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### getItem(index: number)
Gets an inline picture object by its index in the collection.

#### Syntax
```js
inlinePictureCollectionObject.getItem(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|A number that identifies the index location of an inline picture object.|

#### Returns
[InlinePicture](inlinepicture.md)

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
