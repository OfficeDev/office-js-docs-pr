# ChartDataLabels

Represents a collection of all the data labels on a chart point.

## [Properties](#getter-examples)
| Property	   | Type	|Description
|:---------------|:--------|:----------|
|position|string|DataLabelPosition value that represents the position of the data label. Possible values are: None, Center, InsideEnd, InsideBase, OutsideEnd, Left, Right, Top, Bottom, BestFit, Callout.|
|separator|string|String representing the separator used for the data labels on a chart.|
|showBubbleSize|bool|Boolean value representing if the data label bubble size is visible or not.|
|showCategoryName|bool|Boolean value representing if the data label category name is visible or not.|
|showLegendKey|bool|Boolean value representing if the data label legend key is visible or not.|
|showPercentage|bool|Boolean value representing if the data label percentage is visible or not.|
|showSeriesName|bool|Boolean value representing if the data label series name is visible or not.|
|showValue|bool|Boolean value representing if the data label value is visible or not.|

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|format|[ChartDataLabelFormat](chartdatalabelformat.md)|Represents the format of chart data labels, which includes fill and font formatting. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## API Specification

### load(param: object)
Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.

#### Syntax
```js
object.load(param);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|param|object|Optional. Accepts parameter and relationship names as delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns
void

#### Examples
```js

```

[Back](#methods)

### Getter Examples
Make Series Name shown in Datalabels and set the `position` of datalabels to be "top";
```js
var ctx = new Excel.RequestContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	

chart.datalabels.visible = true;
chart.datalabels.position = "top";
chart.datalabels.ShowSeriesName = true;

ctx.executeAsync().then(function () {
		Console.log("Datalabels Shown");
});
```

[Back](#properties)
