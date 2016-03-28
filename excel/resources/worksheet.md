# Worksheet object (JavaScript API for Excel)

An Excel worksheet is a grid of cells. It can contain data, tables, charts, etc.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|id|string|Returns a value that uniquely identifies the worksheet in a given workbook. The value of the identifier remains the same even when the worksheet is renamed or moved. Read-only.|
|name|string|The display name of the worksheet.|
|position|int|The zero-based position of the worksheet within the workbook.|
|visibility|string|The Visibility of the worksheet. Possible values are: Visible, Hidden, VeryHidden.|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|charts|[ChartCollection](chartcollection.md)|Returns collection of charts that are part of the worksheet. Read-only.|
|protection|[WorksheetProtection](worksheetprotection.md)|Returns sheet protection object for a worksheet. Read-only.|
|tables|[TableCollection](tablecollection.md)|Collection of tables that are part of the worksheet. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[activate()](#activate)|void|Activate the worksheet in the Excel UI.|
|[delete()](#delete)|void|Deletes the worksheet from the workbook.|
|[getCell(row: number, column: number)](#getcellrow-number-column-number)|[Range](range.md)|Gets the range object containing the single cell based on row and column numbers. The cell can be outside the bounds of its parent range as long as it stays within the worksheet grid.|
|[getRange(address: string)](#getrangeaddress-string)|[Range](range.md)|Gets the range object specified by the address or name.|
|[getUsedRange(valuesOnly: bool)](#getusedrangevaluesonly-bool)|[Range](range.md)|The used range is the smallest range that encompasses any cells that have a value or formatting assigned to them. If the worksheet is blank, this function will return the top left cell.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in the JavaScript layer, with property and object values specified in the parameter.|

## Method Details


### activate()
Activate the worksheet in the Excel UI.

#### Syntax
```js
worksheetObject.activate();
```

#### Parameters
None

#### Returns
void

#### Examples

```js
Excel.run(function (ctx) { 
	var wSheetName = 'Sheet1';
	var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
	worksheet.activate();
	return ctx.sync(); 
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### delete()
Deletes the worksheet from the workbook.

#### Syntax
```js
worksheetObject.delete();
```

#### Parameters
None

#### Returns
void

#### Examples

```js
Excel.run(function (ctx) { 
	var wSheetName = 'Sheet1';
	var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
	worksheet.delete();
	return ctx.sync(); 
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### getCell(row: number, column: number)
Gets the range object containing the single cell based on row and column numbers. The cell can be outside the bounds of its parent range as long as it stays within the worksheet grid.

#### Syntax
```js
worksheetObject.getCell(row, column);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|row|number|The row number of the cell to be retrieved. Zero-indexed.|
|column|number|The column number of the cell to be retrieved. Zero-indexed.|

#### Returns
[Range](range.md)

#### Examples

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:F8";
	var worksheet = ctx.workbook.worksheets.getItem(sheetName);
	var cell = worksheet.getCell(0,0);
	cell.load('address');
	return ctx.sync().then(function() {
		console.log(cell.address);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### getRange(address: string)
Gets the range object specified by the address or name.

#### Syntax
```js
worksheetObject.getRange(address);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|address|string|Optional. The address or the name of the range. If not specified, the entire worksheet range is returned.|

#### Returns
[Range](range.md)

#### Examples
This example uses range address to get the range object.

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:F8";
	var worksheet = ctx.workbook.worksheets.getItem(sheetName);
	var range = worksheet.getRange(rangeAddress);
	range.load('cellCount');
	return ctx.sync().then(function() {
		console.log(range.cellCount);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

This example uses a named-range to get the range object.

```js

Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeName = 'MyRange';
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeName);
	range.load('address');
	return ctx.sync().then(function() {
		console.log(range.address);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getUsedRange(valuesOnly: bool)
The used range is the smallest range that encompasses any cells that have a value or formatting assigned to them. If the worksheet is blank, this function will return the top left cell.

#### Syntax
```js
worksheetObject.getUsedRange(valuesOnly);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|valuesOnly|bool|Optional. Considers only cells with values as used cells (ignores formatting).|

#### Returns
[Range](range.md)

#### Examples

```js
Excel.run(function (ctx) { 
	var wSheetName = 'Sheet1';
	var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
	var usedRange = worksheet.getUsedRange();
	usedRange.load('address');
	return ctx.sync().then(function() {
			console.log(usedRange.address);
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

Get worksheet properties based on the sheet name.

```js
Excel.run(function (ctx) { 
	var wSheetName = 'Sheet1';
	var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
	worksheet.load('position')
	return ctx.sync().then(function() {
			console.log(worksheet.position);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

Set worksheet position. 

```js
Excel.run(function (ctx) { 
	var wSheetName = 'Sheet1';
	var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
	worksheet.position = 2;
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

