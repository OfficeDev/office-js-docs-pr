# TableCollection Object (JavaScript API for Word)

_Word 2016, Word for iPad, Word for Mac_

Contains the collection of the document's Table objects.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|items|[Table[]](table.md)|A collection of table objects. Read-only.|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|first|[Table](table.md)|Gets the first table in this collection. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getItem(index: number)](#getitemindex-number)|[Table](table.md)|Gets a table object by its index in the collection.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### getItem(index: number)
Gets a table object by its index in the collection.

#### Syntax
```js
tableCollectionObject.getItem(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|A number that identifies the index location of a table object.|

#### Returns
[Table](table.md)

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
