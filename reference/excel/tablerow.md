# TableRow object (JavaScript API for Excel)

Represents a row in a table.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|index|int|Returns the index number of the row within the rows collection of the table. Zero-indexed. Read-only.|
|values|object[][]|Represents the raw values of the specified range. The data returned could be of type string, number, or boolean. A cell that contains an error returns an error string.|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[delete()](#delete)|void|Deletes the row from the table.|
|[getRange()](#getrange)|[Range](range.md)|Returns the range object associated with the entire row.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in the JavaScript layer, with property and object values specified in the parameter.|

## Method Details


### delete()
Deletes the row from the table.

#### Syntax
```js
tableRowObject.delete();
```

#### Parameters
None

#### Returns
void

#### Examples

```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var row = ctx.workbook.tables.getItem(tableName).tableRows.getItemAt(2);
	row.delete();
	return ctx.sync(); 
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### getRange()
Returns the range object associated with the entire row.

#### Syntax
```js
tableRowObject.getRange();
```

#### Parameters
None

#### Returns
[Range](range.md)

#### Examples

```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var row = ctx.workbook.tables.getItem(tableName).tableRows.getItemAt(0);
	var rowRange = row.getRange();
	rowRange.load('address');
	return ctx.sync().then(function() {
		console.log(rowRange.address);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### load(param: object)
Fills the proxy object created in the JavaScript layer, with property and object values specified in the parameter.

#### Syntax
```js
object.load(param);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|param|object|Optional. Accepts parameter and relationship names as a delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns
void
### Property access examples

```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var row = ctx.workbook.tables.getItem(tableName).tableRows.getItem(0);
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

```js
Excel.run(function (ctx) { 
	var tables = ctx.workbook.tables;
	var newValues = [["New", "Values", "For", "New", "Row"]];
	var row = ctx.workbook.tables.getItem(tableName).tableRows.getItemAt(2);
	row.values = newValues;
	row.load('values');
	return ctx.sync().then(function() {
		console.log(row.values);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
