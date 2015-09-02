# TableColumnCollection

Represents a collection of all the columns that are part of the table.

## [Properties](#getter-examples)
| Property	   | Type	|Description
|:---------------|:--------|:----------|
|count|int|Returns the number of columns in the table. Read-only.|
|items|[TableColumn[]](tablecolumn.md)|A collection of tableColumn objects. Read-only.|

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[add(index: number, values: (boolean or string or number)[][])](#addindex-number-values-boolean-or-string-or-number)|[TableColumn](tablecolumn.md)|Adds a new column to the table.|
|[getItem(key: number or string)](#getitemkey-number-or-string)|[TableColumn](tablecolumn.md)|Gets a column object by Name or ID.|
|[getItemAt(index: number)](#getitematindex-number)|[TableColumn](tablecolumn.md)|Gets a column based on its position in the collection.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## API Specification

### add(index: number, values: (boolean or string or number)[][])
Adds a new column to the table.

#### Syntax
```js
tableColumnCollectionObject.add(index, values);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|Specifies the relative position of the new column. The previous column at this position is shifted to the right. The index value should be equal to or less than the last column's index value, so it cannot be used to append a column at the end of the table. Zero-indexed.|
|values|(boolean or string or number)[][]|Optional. A 2-dimensional array of unformatted values of the table column.|

#### Returns
[TableColumn](tablecolumn.md)

#### Examples

```js
var ctx = new Excel.RequestContext();
var tables = ctx.workbook.tables;
var values = [["Sample"], ["Values"], ["For"], ["New"], ["Column"]];
var column = tables.getItem("Table1").columns.add(null, values);
column.load(name);
ctx.executeAsync().then(function () {
	Console.log(column.name);
});
```


[Back](#methods)

### getItem(key: number or string)
Gets a column object by Name or ID.

#### Syntax
```js
tableColumnCollectionObject.getItem(key);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|key|number or string| Column Name or ID.|

#### Returns
[TableColumn](tablecolumn.md)

#### Examples

```js
var ctx = new Excel.RequestContext();
var tablecolumn = ctx.workbook.tables.getItem['Table1'].columns.getItem(0);
tablecolumn.load(name);
ctx.executeAsync().then(function () {
		Console.log(tablecolumn.name);
});
```

```js
var ctx = new Excel.RequestContext();
var tablecolumn = ctx.workbook.tables.getItem['Table1'].columns.getItemAt(0);
tablecolumn.load(name);
ctx.executeAsync().then(function () {
		Console.log(tablecolumn.name);
});
```

[Back](#methods)

### getItemAt(index: number)
Gets a column based on its position in the collection.

#### Syntax
```js
tableColumnCollectionObject.getItemAt(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|Index value of the object to be retrieved. Zero-indexed.|

#### Returns
[TableColumn](tablecolumn.md)

#### Examples
```js
var ctx = new Excel.RequestContext();
var tablecolumn = ctx.workbook.tables.getItem['Table1'].columns.getItemAt(0);
tablecolumn.load(name);
ctx.executeAsync().then(function () {
		Console.log(tablecolumn.name);
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
var tablecolumns = ctx.workbook.tables.getItem['Table1'].columns;
tablecolumns.load(items);
ctx.executeAsync().then(function () {
	Console.log("tablecolumns Count: " + tablecolumns.count);
	for (var i = 0; i < tablecolumns.items.length; i++)
	{
		Console.log(tablecolumns.items[i].name);
	}
});
```
[Back](#properties)
