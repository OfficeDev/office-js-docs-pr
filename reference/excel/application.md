# Application Object (JavaScript API for Excel)

Represents the Excel application that manages the workbook.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|calculationMode|string|Returns the calculation mode used in the workbook. Read-only. Possible values are: `Automatic` Excel controls recalculation,`AutomaticExceptTables` Excel controls recalculation but ignores changes in tables.,`Manual` Calculation is done when the user requests it.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[calculate(calculationType: string)](#calculatecalculationtype-string)|void|Recalculate all currently opened workbooks in Excel.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[suspendApiCalculationUntilNextSync()](#suspendapicalculationuntilnextsync)|void|Suspends calculation until the next "context.sync()" is called. Once set, it is the developer's responsibility to re-calc the workbook, to ensure that any dependencies are propagated.|[1.6](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### calculate(calculationType: string)
Recalculate all currently opened workbooks in Excel.

#### Syntax
```js
applicationObject.calculate(calculationType);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|calculationType|string|Specifies the calculation type to use. Possible values are: `Recalculate` Default-option. Performs normal calculation by calculating all the formulas in the workbook,`Full` Forces a full calculation of the data,`FullRebuild`  Forces a full calculation of the data and rebuilds the dependencies.|

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

### suspendApiCalculationUntilNextSync()
Suspends calculation until the next "context.sync()" is called. Once set, it is the developer's responsibility to re-calc the workbook, to ensure that any dependencies are propagated.

#### Syntax
```js
applicationObject.suspendApiCalculationUntilNextSync();
```

#### Parameters
None

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

