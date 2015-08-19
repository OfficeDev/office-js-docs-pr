# TableRowCollection

Represents a collection of all the rows that are part of the table.

## [Properties](#getter-examples)
| Property	   | Type	|Description
|:---------------|:--------|:----------|
|count|int|Returns the number of rows in the table. Read-only.|
|items|[TableRow[]](tablerow.md)|A collection of tableRow objects. Read-only.|

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[add(index: number, values: (boolean or string or number)[][])](#addindex-number-values-boolean-or-string-or-number)|[TableRow](tablerow.md)|Adds a new row to the table.|
|[getItemAt(index: number)](#getitematindex-number)|[TableRow](tablerow.md)|Gets a row based on its position in the collection.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## API Specification

### add(index: number, values: (boolean or string or number)[][])
Adds a new row to the table.

#### Syntax
```js
tableRowCollectionObject.add(index, values);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|Optional. Specifies the relative position of the new row. If null, the addition happens at the end. Any rows below the inserted row are shifted downwards. Zero-indexed.|
|values|(boolean or string or number)[][]|Optional. A 2-dimensional array of unformatted values of the table row.|

#### Returns
[TableRow](tablerow.md)

#### Examples

```js
var ctx = new Excel.RequestContext();
var tables = ctx.workbook.tables;
var values = [["Sample", "Values", "For", "New", "Row"]];
var row = tables.getItem("Table1").rows.add(null, values);
row.load(index);
ctx.executeAsync().then(function () {
	Console.log(row.index);
});
```

[Back](#methods)

### getItemAt(index: number)
Gets a row based on its position in the collection.

#### Syntax
```js
tableRowCollectionObject.getItemAt(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|Index value of the object to be retrieved. Zero-indexed.|

#### Returns
[TableRow](tablerow.md)

#### Examples

```js
var ctx = new Excel.RequestContext();
var tablerow = ctx.workbook.tables.getItem('Table1').rows.getItemAt(0);
tablerow.load(name);
ctx.executeAsync().then(function () {
		Console.log(tablerow.name);
});
```

[Back](#methods)

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

#### Examples
```js

```

[Back](#methods)

### Getter Examples

```js
var ctx = new Excel.RequestContext();
var tablerows = ctx.workbook.tables.getItem('Table1').rows;
tablerows.load(items);
ctx.executeAsync().then(function () {
	Console.log("tablerows Count: " + tablerows.count);
	for (var i = 0; i < tablerows.items.length; i++)
	{
		Console.log(tablerows.items[i].index);
	}
});
```
[Back](#properties)
