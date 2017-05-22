# TableRowCollection Object (JavaScript API for Excel)

Represents a collection of all the rows that are part of the table.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|count|int|Returns the number of rows in the table. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|items|[TableRow[]](tablerow.md)|A collection of tableRow objects. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[add(index: number, values: object)](#addindex-number-values-object)|[TableRow](tablerow.md)|Adds one or more rows to the table. The return object will be the top of the newly added row(s).|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount()](#getcount)|int|Gets the number of rows in the table.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[TableRow](tablerow.md)|Gets a row based on its position in the collection.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### add(index: number, values: object)
Adds one or more rows to the table. The return object will be the top of the newly added row(s).

#### Syntax
```js
tableRowCollectionObject.add(index, values);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|Optional. Specifies the relative position of the new row. If null or -1, the addition happens at the end. Any rows below the inserted row are shifted downwards. Zero-indexed.|
|values|object|Optional. A 2-dimensional array of unformatted values of the table row.|

#### Returns
[TableRow](tablerow.md)

#### Examples

```js
Excel.run(function (ctx) { 
	var tables = ctx.workbook.tables;
	var values = [["Sample", "Values", "For", "New", "Row"]];
	var row = tables.getItem("Table1").rows.add(null, values);
	row.load('index');
	return ctx.sync().then(function() {
		console.log(row.index);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getCount()
Gets the number of rows in the table.

#### Syntax
```js
tableRowCollectionObject.getCount();
```

#### Parameters
None

#### Returns
int

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
Excel.run(function (ctx) { 
	var tablerow = ctx.workbook.tables.getItem('Table1').rows.getItemAt(0);
	tablerow.load('name');
	return ctx.sync().then(function() {
			console.log(tablerow.name);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
### Property access examples

```js
Excel.run(function (ctx) { 
	var tablerows = ctx.workbook.tables.getItem('Table1').rows;
	tablerows.load('items');
	return ctx.sync().then(function() {
		console.log("tablerows Count: " + tablerows.count);
		for (var i = 0; i < tablerows.items.length; i++)
		{
			console.log(tablerows.items[i].index);
		}
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```