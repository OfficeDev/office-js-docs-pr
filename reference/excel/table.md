# Table object (JavaScript API for Excel)

_Applies to: Excel 2016, Excel Online, Excel for iOS, Office 2016_

Represents an Excel table.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|id|int|Returns a value that uniquely identifies the table in a given workbook. The value of the identifier remains the same even when the table is renamed. Read-only.|
|name|string|Name of the table.|
|showHeaders|bool|Indicates whether the header row is visible or not. This value can be set to show or remove the header row.|
|showTotals|bool|Indicates whether the total row is visible or not. This value can be set to show or remove the total row.|
|style|string|Constant value that represents the Table style. Possible values are: TableStyleLight1 through TableStyleLight21, TableStyleMedium1 through TableStyleMedium28, TableStyleStyleDark1 through TableStyleStyleDark11. A custom, user-defined style present in the workbook can also be specified.|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|columns|[TableColumnCollection](tablecolumncollection.md)|Represents a collection of all the columns in the table. Read-only.|
|rows|[TableRowCollection](tablerowcollection.md)|Represents a collection of all the rows in the table. Read-only.|
|sort|[TableSort](tablesort.md)|Represents the sorting configuration for the table. Read-only.|
|worksheet|[Worksheet](worksheet.md)|The worksheet containing the current table. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[clearFilters()](#clearfilters)|void|Clears all the filters currently applied on the table.|
|[convertToRange()](#converttorange)|[Range](range.md)|Converts the table into a normal range of cells. All data is preserved.|
|[delete()](#delete)|void|Deletes the table.|
|[getDataBodyRange()](#getdatabodyrange)|[Range](range.md)|Gets the range object associated with the data body of the table.|
|[getHeaderRowRange()](#getheaderrowrange)|[Range](range.md)|Gets the range object associated with header row of the table.|
|[getRange()](#getrange)|[Range](range.md)|Gets the range object associated with the entire table.|
|[getTotalRowRange()](#gettotalrowrange)|[Range](range.md)|Gets the range object associated with the totals row of the table.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in the JavaScript layer, with property and object values specified in the parameter.|
|[reapplyFilters()](#reapplyfilters)|void|Reapplies all the filters currently on the table.|

## Method Details


### clearFilters()
Clears all the filters currently applied on the table.

#### Syntax
```js
tableObject.clearFilters();
```

#### Parameters
None

#### Returns
void

### convertToRange()
Converts the table into a normal range of cells. All data is preserved.

#### Syntax
```js
tableObject.convertToRange();
```

#### Parameters
None

#### Returns
[Range](range.md)

#### Examples
```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var table = ctx.workbook.tables.getItem(tableName);
	table.convertToRange();
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### delete()
Deletes the table.

#### Syntax
```js
tableObject.delete();
```

#### Parameters
None

#### Returns
void

#### Examples
```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var table = ctx.workbook.tables.getItem(tableName);
	table.delete();
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### getDataBodyRange()
Gets the range object associated with the data body of the table.

#### Syntax
```js
tableObject.getDataBodyRange();
```

#### Parameters
None

#### Returns
[Range](range.md)

#### Examples
```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var table = ctx.workbook.tables.getItem(tableName);
	var tableDataRange = table.getDataBodyRange();
	tableDataRange.load('address')
	return ctx.sync().then(function() {
			console.log(tableDataRange.address);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getHeaderRowRange()
Gets the range object associated with the header row of the table.

#### Syntax
```js
tableObject.getHeaderRowRange();
```

#### Parameters
None

#### Returns
[Range](range.md)

#### Examples
```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var table = ctx.workbook.tables.getItem(tableName);
	var tableHeaderRange = table.getHeaderRowRange();
	tableHeaderRange.load('address');
	return ctx.sync().then(function() {
		console.log(tableHeaderRange.address);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### getRange()
Gets the range object associated with the entire table.

#### Syntax
```js
tableObject.getRange();
```

#### Parameters
None

#### Returns
[Range](range.md)

#### Examples
```js
Excel.run(function (ctx) { 
	var table = ctx.workbook.tables.getItem(tableName);
	var tableRange = table.getRange();
	tableRange.load('address');	
	return ctx.sync().then(function() {
			console.log(tableRange.address);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### getTotalRowRange()
Gets the range object associated with the totals row of the table.

#### Syntax
```js
tableObject.getTotalRowRange();
```

#### Parameters
None

#### Returns
[Range](range.md)

#### Examples
```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var table = ctx.workbook.tables.getItem(tableName);
	var tableTotalsRange = table.getTotalRowRange();
	tableTotalsRange.load('address');	
	return ctx.sync().then(function() {
			console.log(tableTotalsRange.address);
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

Get a table by name. 

```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var table = ctx.workbook.tables.getItem(tableName);
	table.load('index')
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

Get a table by index.

```js
Excel.run(function (ctx) { 
	var index = 0;
	var table = ctx.workbook.tables.getItemAt(0);
	table.name('name')
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

Set table style. 

```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var table = ctx.workbook.tables.getItem(tableName);
	table.name = 'Table1-Renamed';
	table.showTotals = false;
	table.tableStyle = 'TableStyleMedium2';
	table.load('tableStyle');
	return ctx.sync().then(function() {
			console.log(table.tableStyle);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
