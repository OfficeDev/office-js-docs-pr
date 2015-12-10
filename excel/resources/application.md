# Application object (JavaScript API for Excel)

_Applies to: Excel 2016, Excel Online, Office 2016_

Represents the Excel application that manages the workbook.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|calculationMode|string|Returns the calculation mode used in the workbook. Read-only. Possible values are: `Automatic` Excel controls recalculation, `AutomaticExceptTables` Excel controls recalculation but ignores changes in tables, `Manual` Calculation is done when the user requests it.|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[calculate(calculationType: string)](#calculatecalculationtype-string)|void|Recalculate all currently open workbooks in Excel.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

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
|calculationType|string|Specifies the calculation type to use. Possible values are: `Recalculate` Default-option, Performs normal calculation by calculating all the formulas in the workbook, `Full` Forces a full calculation of the data, `FullRebuild` Forces a full calculation of the data and rebuilds the dependencies.|

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

