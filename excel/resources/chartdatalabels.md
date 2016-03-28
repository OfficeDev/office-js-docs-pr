# ChartDataLabels object (JavaScript API for Excel)

Represents a collection of all the data labels on a chart point.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|position|string|DataLabelPosition value that represents the position of the data label. Possible values are: None, Center, InsideEnd, InsideBase, OutsideEnd, Left, Right, Top, Bottom, BestFit, Callout. Write-only.|
|separator|string|String representing the separator used for the data labels on a chart. Write-only.|
|showBubbleSize|bool|Boolean value representing if the data label bubble size is visible or not. Write-only.|
|showCategoryName|bool|Boolean value representing if the data label category name is visible or not. Write-only.|
|showLegendKey|bool|Boolean value representing if the data label legend key is visible or not. Write-only.|
|showPercentage|bool|Boolean value representing if the data label percentage is visible or not. Write-only.|
|showSeriesName|bool|Boolean value representing if the data label series name is visible or not. Write-only.|
|showValue|bool|Boolean value representing if the data label value is visible or not. Write-only.|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|format|[ChartDataLabelFormat](chartdatalabelformat.md)|Represents the format of chart data labels, which includes fill and font formatting. Read-only.|

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

Make the series name shown in data labels and set the `position` of data labels to be "top".

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	chart.datalabels.visible = true;
	chart.datalabels.position = "top";
	chart.datalabels.ShowSeriesName = true;
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
