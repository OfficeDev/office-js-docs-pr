# TableRowCollection Object (JavaScript API for Word)

_Word 2016, Word for iPad, Word for Mac, Word Online_

Contains the collection of the document's TableRow objects.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[TableRow[]](tablerow.md)|A collection of tableRow objects. Read-only.|[1.3](../requirement-sets/word-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getFirst()](#getfirst)|[TableRow](tablerow.md)|Gets the first row in this collection. Throws if this collection is empty.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[getFirstOrNullObject()](#getfirstornullobject)|[TableRow](tablerow.md)|Gets the first row in this collection. Returns a null object if this collection is empty.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[getItem(index: number)](#getitemindex-number)|[TableRow](tablerow.md)|Gets a table row object by its index in the collection.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[1.1](../requirement-sets/word-api-requirement-sets.md)|

## Method Details


### getFirst()
Gets the first row in this collection. Throws if this collection is empty.

#### Syntax
```js
tableRowCollectionObject.getFirst();
```

#### Parameters
None

#### Returns
[TableRow](tablerow.md)

### getFirstOrNullObject()
Gets the first row in this collection. Returns a null object if this collection is empty.

#### Syntax
```js
tableRowCollectionObject.getFirstOrNullObject();
```

#### Parameters
None

#### Returns
[TableRow](tablerow.md)

### getItem(index: number)
Gets a table row object by its index in the collection.

#### Syntax
```js
tableRowCollectionObject.getItem(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|index|number|A number that identifies the index location of a table row object.|

#### Returns
[TableRow](tablerow.md)

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
