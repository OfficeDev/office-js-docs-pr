# ChartAxis object (JavaScript API for Excel)

_Applies to: Excel 2016, Excel Online, Office 2016_

Represents a single axis in a chart.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|majorUnit|object|Represents the interval between two major tick marks. Can be set to a numeric value or an empty string.  The return value is always a number.|
|maximum|object|Represents the maximum value for the value axis.  Can be set to a numeric value or an empty string (for automatic axis values).  The return value is always a number.|
|minimum|object|Represents the minimum value for the value axis. Can be set to a numeric value or an empty string (for automatic axis values).  The return value is always a number.|
|minorUnit|object|Represents the interval between two minor tick marks. Can be set to a numeric value or an empty string (for automatic axis values). The return value is always a number.|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|format|[ChartAxisFormat](chartaxisformat.md)|Represents the formatting of a chart object, which includes line and font formatting. Read-only.|
|majorGridlines|[ChartGridlines](chartgridlines.md)|Returns a Gridlines object that represents the major gridlines for the specified axis. Read-only.|
|minorGridlines|[ChartGridlines](chartgridlines.md)|Returns a Gridlines object that represents the minor gridlines for the specified axis. Read-only.|
|title|[ChartAxisTitle](chartaxistitle.md)|Represents the axis title. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in the JavaScript layer, with property and object values specified in the parameter.|

## Method Details

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
Get the `maximum` of chart axis from Chart1.

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	var axis = chart.axes.valueaxis;
	axis.load('maximum');
	return ctx.sync().then(function() {
			console.log(axis.maximum);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

Set the  `maximum`,  `minimum`,  `majorunit`,or `minorunit` of value axis. 

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	chart.axes.valueaxis.maximum = 5;
	chart.axes.valueaxis.minimum = 0;
	chart.axes.valueaxis.majorunit = 1;
	chart.axes.valueaxis.minorunit = 0.2;
	return ctx.sync().then(function() {
			console.log("Axis Settings Changed");
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
