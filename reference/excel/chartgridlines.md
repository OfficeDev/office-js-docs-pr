# ChartGridlines Object (JavaScript API for Excel)

Represents major or minor gridlines on a chart axis.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|visible|bool|Boolean value representing if the axis gridlines are visible or not.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|format|[ChartGridlinesFormat](chartgridlinesformat.md)|Represents the formatting of chart gridlines. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## Methods
None


## Method Details

### Property access examples

Get the `visible` of Major Gridlines on value axis of Chart1

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	var majGridlines = chart.axes.valueaxis.majorGridlines;
	majGridlines.load('visible');
	return ctx.sync().then(function() {
			console.log(majGridlines.visible);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

Set to show major gridlines on valueAxis of Chart1

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	chart.axes.valueAxis.majorGridlines.visible = true;
	return ctx.sync().then(function() {
			console.log("Axis Gridlines Added ");
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
