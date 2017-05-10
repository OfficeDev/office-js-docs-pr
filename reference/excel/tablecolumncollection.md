# TableColumnCollection Object (JavaScript API for Excel)

Represents a collection of all the columns that are part of the table.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|count|int|Returns the number of columns in the table. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|items|[TableColumn[]](tablecolumn.md)|A collection of tableColumn objects. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[add(index: number, values: object, name: string)](#addindex-number-values-object-name-string)|[TableColumn](tablecolumn.md)|Adds a new column to the table.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount()](#getcount)|int|Gets the number of columns in the table.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(key: object)](#getitemkey-object)|[TableColumn](tablecolumn.md)|Gets a column object by Name or ID.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[TableColumn](tablecolumn.md)|Gets a column based on its position in the collection.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNullObject(key: object)](#getitemornullobjectkey-object)|[TableColumn](tablecolumn.md)|Gets a column object by Name or ID. If the column does not exist, will return a null object.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### add(index: number, values: object, name: string)
Adds a new column to the table.

#### Syntax
```js
tableColumnCollectionObject.add(index, values, name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|Optional. Specifies the relative position of the new column. If null or -1, the addition happens at the end. Columns with a higher index will be shifted to the side. Zero-indexed.|
|values|object|Optional. A 2-dimensional array of unformatted values of the table column.|
|name|string|Optional. Specifies the name of the new column. If null, the default name will be used.|

#### Returns
[TableColumn](tablecolumn.md)

#### Examples

```js
Excel.run(function (ctx) { 
	var tables = ctx.workbook.tables;
	var values = [["Sample"], ["Values"], ["For"], ["New"], ["Column"]];
	var column = tables.getItem("Table1").columns.add(null, values);
	column.load('name');
	return ctx.sync().then(function() {
		console.log(column.name);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### getCount()
Gets the number of columns in the table.

#### Syntax
```js
tableColumnCollectionObject.getCount();
```

#### Parameters
None

#### Returns
int

### getItem(key: object)
Gets a column object by Name or ID.

#### Syntax
```js
tableColumnCollectionObject.getItem(key);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|key|object| Column Name or ID.|

#### Returns
[TableColumn](tablecolumn.md)

#### Examples

```js
Excel.run(function (ctx) { 
	var tablecolumn = ctx.workbook.tables.getItem('Table1').columns.getItem(0);
	tablecolumn.load('name');
	return ctx.sync().then(function() {
			console.log(tablecolumn.name);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


#### Examples
```js
Excel.run(function (ctx) { 
	var tablecolumn = ctx.workbook.tables.getItem['Table1'].columns.getItemAt(0);
	tablecolumn.load('name');
	return ctx.sync().then(function() {
			console.log(tablecolumn.name);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

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
Excel.run(function (ctx) { 
	var tablecolumn = ctx.workbook.tables.getItem['Table1'].columns.getItemAt(0);
	tablecolumn.load('name');
	return ctx.sync().then(function() {
			console.log(tablecolumn.name);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getItemOrNullObject(key: object)
Gets a column object by Name or ID. If the column does not exist, will return a null object.

#### Syntax
```js
tableColumnCollectionObject.getItemOrNullObject(key);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|key|object| Column Name or ID.|

#### Returns
[TableColumn](tablecolumn.md)
### Property access examples

```js
Excel.run(function (ctx) { 
	var tablecolumns = ctx.workbook.tables.getItem('Table1').columns;
	tablecolumns.load('items');
	return ctx.sync().then(function() {
		console.log("tablecolumns Count: " + tablecolumns.count);
		for (var i = 0; i < tablecolumns.items.length; i++)
		{
			console.log(tablecolumns.items[i].name);
		}
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```