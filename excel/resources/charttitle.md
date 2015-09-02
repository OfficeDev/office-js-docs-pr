# ChartTitle

Represents a chart title object of a chart.

## [Properties](#getter-and-setter-examples)
| Property	   | Type	|Description
|:---------------|:--------|:----------|
|overlay|bool|Boolean value representing if the chart title will overlay the chart or not.|
|text|string|Represents the title text of a chart.|
|visible|bool|A boolean value the represents the visibility of a chart title object.|

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|format|[ChartTitleFormat](charttitleformat.md)|Represents the formatting of a chart title, which includes fill and font formatting. Read-only.|

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

### Getter and Setter Examples

Get the `text` of Chart Title from Chart1.

```js
var ctx = new Excel.RequestContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	

var title = chart.title;
title.load(text);
ctx.executeAsync().then(function () {
		Console.log(title.text);
});
```

Set the `text` of Chart Title to "My Chart" and Make it show on top of the chart without overlaying.

```js
var ctx = new Excel.RequestContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	

chart.title.text= "My Chart"; 
chart.title.visible=true;
chart.title.overlay=true;

ctx.executeAsync().then(function () {
		Console.log("Char Title Changed");
});
```

[Back](#properties)
