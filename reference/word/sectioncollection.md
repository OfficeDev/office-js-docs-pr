# SectionCollection Object (JavaScript API for Word)

_Word 2016, Word for iPad, Word for Mac, Word Online_

Contains the collection of the document's [section](section.md) objects.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[Section[]](section.md)|A collection of section objects. Read-only.|[1.1](../requirement-sets/word-api-requirement-sets.md)|

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getFirst()](#getfirst)|[Section](section.md)|Gets the first section in this collection. Throws if this collection is empty.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[getFirstOrNullObject()](#getfirstornullobject)|[Section](section.md)|Gets the first section in this collection. Returns a null object if this collection is empty.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[getItem(index: number)](#getitemindex-number)|[Section](section.md)|Gets a section object by its index in the collection.|[1.1](../requirement-sets/word-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[1.1](../requirement-sets/word-api-requirement-sets.md)|

## Method Details


### getFirst()
Gets the first section in this collection. Throws if this collection is empty.

#### Syntax
```js
sectionCollectionObject.getFirst();
```

#### Parameters
None

#### Returns
[Section](section.md)

### getFirstOrNullObject()
Gets the first section in this collection. Returns a null object if this collection is empty.

#### Syntax
```js
sectionCollectionObject.getFirstOrNullObject();
```

#### Parameters
None

#### Returns
[Section](section.md)

### getItem(index: number)
Gets a section object by its index in the collection.

#### Syntax
```js
sectionCollectionObject.getItem(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|index|number|A number that identifies the index location of a section object.|

#### Returns
[Section](section.md)

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
