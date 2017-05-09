# InlinePictureCollection Object (JavaScript API for Word)

_Word 2016, Word for iPad, Word for Mac, Word Online_

Contains a collection of [inlinePicture](inlinePicture.md) objects.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[InlinePicture[]](inlinepicture.md)|A collection of inlinePicture objects. Read-only.|[1.1](../requirement-sets/word-api-requirement-sets.md)|

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getFirst()](#getfirst)|[InlinePicture](inlinepicture.md)|Gets the first inline image in this collection. Throws if this collection is empty.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[getFirstOrNullObject()](#getfirstornullobject)|[InlinePicture](inlinepicture.md)|Gets the first inline image in this collection. Returns a null object if this collection is empty.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[getItem(index: number)](#getitemindex-number)|[InlinePicture](inlinepicture.md)|Gets an inline picture object by its index in the collection.|[1.1](../requirement-sets/word-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[1.1](../requirement-sets/word-api-requirement-sets.md)|

## Method Details


### getFirst()
Gets the first inline image in this collection. Throws if this collection is empty.

#### Syntax
```js
inlinePictureCollectionObject.getFirst();
```

#### Parameters
None

#### Returns
[InlinePicture](inlinepicture.md)

### getFirstOrNullObject()
Gets the first inline image in this collection. Returns a null object if this collection is empty.

#### Syntax
```js
inlinePictureCollectionObject.getFirstOrNullObject();
```

#### Parameters
None

#### Returns
[InlinePicture](inlinepicture.md)

### getItem(index: number)
Gets an inline picture object by its index in the collection.

#### Syntax
```js
inlinePictureCollectionObject.getItem(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
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
|:---------------|:--------|:----------|:---|
|param|object|Optional. Accepts parameter and relationship names as delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns
void
