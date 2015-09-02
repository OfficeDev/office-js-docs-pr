# ChartGridlines

Represents major or minor gridlines on a chart axis.

## [Properties](#getter-and-setter-examples)
| Property	   | Type	|Description
|:---------------|:--------|:----------|
|visible|bool|Boolean value representing if the axis gridlines are visible or not.|

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|format|[ChartGridlinesFormat](chartgridlinesformat.md)|Represents the formatting of chart gridlines. Read-only.|

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

Get the `visible` of Major Gridlines on value axis of Chart1
```js
var ctx = new Excel.RequestContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	

var majGridlines = chart.axes.valueaxis.majorGridlines;
majGridlines.load(visible);
ctx.executeAsync().then(function () {
		Console.log(majGridlines.visible);
});
```

Set to show major gridlines on valueAxis of Chart1

```js
var ctx = new Excel.RequestContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	

chart.axes.valueaxis.majorgridlines.visible = true;

ctx.executeAsync().then(function () {
		Console.log("Axis Gridlines Added ");
});
```

[Back](#properties)
