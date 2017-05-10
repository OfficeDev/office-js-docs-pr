# WorksheetCollection Object (JavaScript API for Excel)

Represents a collection of worksheet objects that are part of the workbook.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[Worksheet[]](worksheet.md)|A collection of worksheet objects. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[add(name: string)](#addname-string)|[Worksheet](worksheet.md)|Adds a new worksheet to the workbook. The worksheet will be added at the end of existing worksheets. If you wish to activate the newly added worksheet, call ".activate() on it.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getActiveWorksheet()](#getactiveworksheet)|[Worksheet](worksheet.md)|Gets the currently active worksheet in the workbook.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount(visibleOnly: bool)](#getcountvisibleonly-bool)|int|Gets the number of worksheets in the collection.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getFirst(visibleOnly: bool)](#getfirstvisibleonly-bool)|[Worksheet](worksheet.md)|Gets the first worksheet in the collection.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(key: string)](#getitemkey-string)|[Worksheet](worksheet.md)|Gets a worksheet object using its Name or ID.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNullObject(key: string)](#getitemornullobjectkey-string)|[Worksheet](worksheet.md)|Gets a worksheet object using its Name or ID. If the worksheet does not exist, will return a null object.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getLast(visibleOnly: bool)](#getlastvisibleonly-bool)|[Worksheet](worksheet.md)|Gets the last worksheet in the collection.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### add(name: string)
Adds a new worksheet to the workbook. The worksheet will be added at the end of existing worksheets. If you wish to activate the newly added worksheet, call ".activate() on it.

#### Syntax
```js
worksheetCollectionObject.add(name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|name|string|Optional. The name of the worksheet to be added. If specified, name should be unqiue. If not specified, Excel determines the name of the new worksheet.|

#### Returns
[Worksheet](worksheet.md)

#### Examples

```js
Excel.run(function (ctx) { 
	var wSheetName = 'Sample Name';
	var worksheet = ctx.workbook.worksheets.add(wSheetName);
	worksheet.load('name');
	return ctx.sync().then(function() {
		console.log(worksheet.name);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### getActiveWorksheet()
Gets the currently active worksheet in the workbook.

#### Syntax
```js
worksheetCollectionObject.getActiveWorksheet();
```

#### Parameters
None

#### Returns
[Worksheet](worksheet.md)

#### Examples

```js
Excel.run(function (ctx) {  
	var activeWorksheet = ctx.workbook.worksheets.getActiveWorksheet();
	activeWorksheet.load('name');
	return ctx.sync().then(function() {
			console.log(activeWorksheet.name);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### getCount(visibleOnly: bool)
Gets the number of worksheets in the collection.

#### Syntax
```js
worksheetCollectionObject.getCount(visibleOnly);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|visibleOnly|bool|Optional. Considers only the visible cells.|

#### Returns
int

### getFirst(visibleOnly: bool)
Gets the first worksheet in the collection.

#### Syntax
```js
worksheetCollectionObject.getFirst(visibleOnly);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|visibleOnly|bool|Optional. If true, considers only visible worksheets, skipping over any hidden ones.|

#### Returns
[Worksheet](worksheet.md)

### getItem(key: string)
Gets a worksheet object using its Name or ID.

#### Syntax
```js
worksheetCollectionObject.getItem(key);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|key|string|The Name or ID of the worksheet.|

#### Returns
[Worksheet](worksheet.md)

### getItemOrNullObject(key: string)
Gets a worksheet object using its Name or ID. If the worksheet does not exist, will return a null object.

#### Syntax
```js
worksheetCollectionObject.getItemOrNullObject(key);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|key|string|The Name or ID of the worksheet.|

#### Returns
[Worksheet](worksheet.md)

### getLast(visibleOnly: bool)
Gets the last worksheet in the collection.

#### Syntax
```js
worksheetCollectionObject.getLast(visibleOnly);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|visibleOnly|bool|Optional. If true, considers only visible worksheets, skipping over any hidden ones.|

#### Returns
[Worksheet](worksheet.md)
### Property access examples
```js
Excel.run(function (ctx) { 
	var worksheets = ctx.workbook.worksheets;
	worksheets.load('items');
	return ctx.sync().then(function() {
		for (var i = 0; i < worksheets.items.length; i++)
		{
			console.log(worksheets.items[i].name);
			console.log(worksheets.items[i].index);
		}
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
