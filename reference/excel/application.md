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
|[suspendCalculationUntilNextSync()](#suspendcalculationuntilnextsync)|void|Suspends calculation until the next "context.sync()" is called. Once set, it is the developer's responsibility to re-calc the workbook, to ensure that any dependencies are propagated.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### calculate(calculationType: string)
Recalculate all currently opened workbooks in Excel.

#### Syntax
```js
applicationObject.calculate(calculationType);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|calculationType|string|Specifies the calculation type to use. Possible values are: `Recalculate` Recalculates all cells that Excel has marked as dirty, that is, dependents of volatile or changed data, and cells programmatically marked as dirty. `Full` This will mark all cells as dirty and then recalculate them. `FullRebuild` This will force a rebuild of the entire calculation chain, mark all cells as dirty and then recalculate all cells.|

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

### suspendCalculationUntilNextSync()
Suspends calculation until the next "context.sync()" is called. Once set, it is the developer's responsibility to re-calc the workbook, to ensure that any dependencies are propagated.

#### Syntax
```js
applicationObject.suspendCalculationUntilNextSync();
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

