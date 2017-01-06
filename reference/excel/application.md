# Application object (JavaScript API for Excel)

Represents the Excel application that manages the workbook.

## Properties

| Property	   | Type	|Description|Req. Set|
|:---------------|:--------|:----------|:----------|
|calculationMode|string|Returns the calculation mode used in the workbook. Read-only. Possible values are: `Automatic` Excel controls recalculation, `AutomaticExceptTables` Excel controls recalculation but ignores changes in tables, `Manual` Calculation is done when the user requests it.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description|Req. Set|
|:---------------|:--------|:----------|:----------|
|[calculate(calculationType: string)](#calculatecalculationtype-string)|void|Recalculate all currently open workbooks in Excel.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### calculate(calculationType: string)
Recalculate all currently open workbooks in Excel.

#### Syntax
```js
applicationObject.calculate(calculationType);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|calculationType|string|Specifies the calculation type to use. Possible values are: `Recalculate` This is a soft recalculation and is mainly for backwards compatibilty. `Full` Recalculates all cells that Excel has marked as dirty, that is, dependents of volatile or changed data, and cells programmatically marked as dirty. `FullRebuild` Recalculates all cells in all open workbooks.|

#### Returns
void

#### Examples
```js
Excel.run(function (ctx) { 
	ctx.workbook.application.calculate('Full');
	return ctx.sync(); 
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
|param|object|Optional. Accepts parameter and relationship names as a delimited string or an array. Or, accepts a [loadOption](loadoption.md) object.|

#### Returns
void
### Property access examples
```js
Excel.run(function (ctx) { 
	var application = ctx.workbook.application;
	application.load('calculationMode');
	return ctx.sync().then(function() {
		console.log(application.calculationMode);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
