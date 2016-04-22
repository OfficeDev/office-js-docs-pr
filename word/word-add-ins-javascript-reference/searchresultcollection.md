# SearchResultCollection Object (JavaScript API for Word)

_Word 2016, Word for iPad, Word for Mac_

Contains a collection of [range](range.md) objects as a result of a search operation.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|items|[SearchResult[]](searchresult.md)|A collection of searchResult objects. Read-only.|

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|first|[Range](range.md)|Gets the first searched result in this collection. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getItem(index: number)](#getitemindex-number)|[Range](range.md)|Gets a range object by its index in the collection.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### getItem(index: number)
Gets a range object by its index in the collection.

#### Syntax
```js
searchResultCollectionObject.getItem(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number| A number that identifies the index location of a range object. |

#### Returns
[Range](range.md)

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
