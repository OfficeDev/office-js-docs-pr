# TableCollection object (JavaScript API for Excel)

_Applies to: Excel 2016, Excel Online, Office 2016_

Represents a collection of all the tables that are part of the workbook.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|count|int|Returns the number of tables in the workbook. Read-only.|
|items|[Table[]](table.md)|A collection of table objects. Read-only.|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[add(address: string, hasHeaders: bool)](#addaddress-string-hasheaders-bool)|[Table](table.md)|Create a new table. The range source address determines the worksheet under which the table will be added. If the table can't be added (e.g., because the address is invalid, or the table would overlap with another table), an error is thrown.|
|[getItem(key: number or string)](#getitemkey-number-or-string)|[Table](table.md)|Gets a table by name or ID.|
|[getItemAt(index: number)](#getitematindex-number)|[Table](table.md)|Gets a table based on its position in the collection.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in the JavaScript layer, with property and object values specified in the parameter.|

## Method Details

### add(address: string, hasHeaders: bool)
Creates a new table. The range source address determines the worksheet under which the table will be added. If the table can't be added (e.g., because the address is invalid, or the table would overlap with another table), an error is thrown.

#### Syntax
```js
tableCollectionObject.add(address, hasHeaders);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|address|string|Address or name of the range object representing the data source. If the address does not contain a sheet name, the currently active sheet is used.|
|hasHeaders|bool|Boolean value that indicates whether the data being imported has column labels. If the source does not contain headers (i.e., when this property is set to false), Excel will automatically generate a header, shifting the data down by one row.|

#### Returns
[Table](table.md)

#### Examples

```js
Excel.run(function (ctx) { 
	var table = ctx.workbook.tables.add('Sheet1!A1:E7', true);
	table.load('name');
	return ctx.sync().then(function() {
		console.log(table.name);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
### getItem(key: number or string)
Gets a table by name or ID.

#### Syntax
```js
tableCollectionObject.getItem(key);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|key|number or string|Name or ID of the table to be retrieved.|

#### Returns
[Table](table.md)

#### Examples

```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var table = ctx.workbook.tables.getItem(tableName);
	return ctx.sync().then(function() {
			console.log(table.index);
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
	var table = ctx.workbook.tables.getItemAt(0);
	return ctx.sync().then(function() {
			console.log(table.name);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getItemAt(index: number)
Gets a table based on its position in the collection.

#### Syntax
```js
tableCollectionObject.getItemAt(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|Index value of the object to be retrieved. Zero-indexed.|

#### Returns
[Table](table.md)

#### Examples

```js
Excel.run(function (ctx) { 
	var table = ctx.workbook.tables.getItemAt(0);
	return ctx.sync().then(function() {
			console.log(table.name);
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
	var tables = ctx.workbook.tables;
	tables.load('items');
	return ctx.sync().then(function() {
		console.log("tables Count: " + tables.count);
		for (var i = 0; i < tables.items.length; i++)
		{
			console.log(tables.items[i].name);
		}
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

Get the number of tables.

```js
Excel.run(function (ctx) { 
	var tables = ctx.workbook.tables;
	tables.load('count');
	return ctx.sync().then(function() {
		console.log(tables.count);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```