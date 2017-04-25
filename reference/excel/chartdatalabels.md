# ChartDataLabels Object (JavaScript API for Excel)

Represents a collection of all the data labels on a chart point.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|position|string|DataLabelPosition value that represents the position of the data label. Possible values are: None, Center, InsideEnd, InsideBase, OutsideEnd, Left, Right, Top, Bottom, BestFit, Callout.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|separator|string|String representing the separator used for the data labels on a chart.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showBubbleSize|bool|Boolean value representing if the data label bubble size is visible or not.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showCategoryName|bool|Boolean value representing if the data label category name is visible or not.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showLegendKey|bool|Boolean value representing if the data label legend key is visible or not.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showPercentage|bool|Boolean value representing if the data label percentage is visible or not.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showSeriesName|bool|Boolean value representing if the data label series name is visible or not.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showValue|bool|Boolean value representing if the data label value is visible or not.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|format|[ChartDataLabelFormat](chartdatalabelformat.md)|Represents the format of chart data labels, which includes fill and font formatting. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## Methods
None


## Method Details

### Property access examples

Make Series Name shown in Datalabels and set the `position` of datalabels to be "top";

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	chart.datalabels.showValue = true;
	chart.datalabels.position = "top";
	chart.datalabels.showSeriesName = true;
	return ctx.sync().then(function() {
			console.log("Datalabels Shown");
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
