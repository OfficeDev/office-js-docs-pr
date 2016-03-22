# ChartLineFormat object (JavaScript API for Excel)

Enapsulates the formatting options for line elements.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|color|string|HTML color code representing the color of lines in the chart.|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[clear()](#clear)|void|Clear the line format of a chart element.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in the JavaScript layer, with property and object values specified in the parameter.|

## Method Details


### clear()
Clear the line format of a chart element.

#### Syntax
```js
chartLineFormatObject.clear();
```

#### Parameters
None

#### Returns
void

#### Examples

Clear the line format of the major gridlines for the value axis of the chart named "Chart1".

```js
Excel.run(function (ctx) { 
	var gridlines = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").axes.valueaxis.majorGridlines;	
	gridlines.format.line.clear();
	return ctx.sync().then(function() {
			console.log("Chart Major Gridlines Format Cleared");
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

Set chart major gridlines for the value axis to red.

```js
Excel.run(function (ctx) { 
	var gridlines = ctx.workbook.worksheets.getItem("Sheet1").charts.axes.valueaxis.majorGridlines;
	gridlines.format.line.color = "#FF0000";
	return ctx.sync().then(function() {
			console.log("Chart Gridlines Color Updated");
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
