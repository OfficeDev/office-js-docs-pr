# TableCellCollection Object (JavaScript API for OneNote)

_Applies to: OneNote Online_  

Contains a collection of TableCell objects.

To provide feedback on this API, you can [file an issue in GitHub](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCellCollection).

## Properties

| Property	   | Type	|Description|
|:---------------|:--------|:----------|
|count|int|Returns the number of table cells in this collection. Read-only.|
|items|[TableCell[]](tablecell.md)|A collection of TableCell objects. Read-only.|

_See [property access examples](#property-access-examples)_.

## Relationships

None


## Methods

| Method		   | Return Type	|Description| 
|:---------------|:--------|:----------|
|[getItem(index: number or string)](#getitemindex-number-or-string)|[TableCell](tablecell.md)|Gets a table cell object by ID or by its index in the collection. Read-only.|
|[getItemAt(index: number)](#getitematindex-number)|[TableCell](tablecell.md)|Gets a table cell at its position in the collection.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in the JavaScript layer with property and object values specified in the parameter.|

## Method details

### getItem(index: number or string)

Gets a table cell object by ID or by its index in the collection. Read-only.

#### Syntax

```js
tableCellCollectionObject.getItem(index);
```

#### Parameters

| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number or string|A number that identifies the index location of a TableCell object.|

#### Returns

[TableCell](tablecell.md)

<br/>

### getItemAt(index: number)

Gets a table cell at its position in the collection.
 
#### Syntax

```js
tableCellCollectionObject.getItemAt(index);
```

#### Parameters

| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|Index value of the object to be retrieved. Zero-indexed.|

#### Returns

[TableCell](tablecell.md)

<br/>

### load(param: object)

Fills the proxy object created in the JavaScript layer with property and object values specified in the parameter.

#### Syntax

```js
object.load(param);
```

#### Parameters

| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|param|object|Optional. Accepts parameter and relationship names as a delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns

Void

<br/>
