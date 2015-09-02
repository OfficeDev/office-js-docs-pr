# ChartAxis

Represents a single axis in a chart.

## [Properties](#getter-and-setter-examples)
| Property	   | Type	|Description
|:---------------|:--------|:----------|
|majorUnit|object|Represents the interval between two major tick marks. Can be set to a numeric value or an empty string.  The returned value is always a number.|
|maximum|object|Represents the maximum value on the value axis.  Can be set to a numeric value or an empty string (for automatic axis values).  The returned value is always a number.|
|minimum|object|Represents the minimum value on the value axis. Can be set to a numeric value or an empty string (for automatic axis values).  The returned value is always a number.|
|minorUnit|object|Represents the interval between two minor tick marks. "Can be set to a numeric value or an empty string (for automatic axis values). The returned value is always a number.|

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|format|[ChartAxisFormat](chartaxisformat.md)|Represents the formatting of a chart object, which includes line and font formatting. Read-only.|
|majorGridlines|[ChartGridlines](chartgridlines.md)|Returns a gridlines object that represents the major gridlines for the specified axis. Read-only.|
|minorGridlines|[ChartGridlines](chartgridlines.md)|Returns a Gridlines object that represents the minor gridlines for the specified axis. Read-only.|
|title|[ChartAxisTitle](chartaxistitle.md)|Represents the axis title. Read-only.|

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
Get the `maximum` of Chart Axis from Chart1
```js
var ctx = new Excel.RequestContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	

var axis = chart.axes.valueaxis;
axis.load(maximum);
ctx.executeAsync().then(function () {
		Console.log(axis.maximum);
});
```

Set the  `maximum`,  `minimum`,  `majorunit`, `minorunit` of valueaxis. 

```js
var ctx = new Excel.RequestContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	

chart.axes.valueaxis.maximum = 5;
chart.axes.valueaxis.minimum = 0;
chart.axes.valueaxis.majorunit = 1;
chart.axes.valueaxis.minorunit = 0.2;

ctx.executeAsync().then(function () {
		Console.log("Axis Settings Changed");
});
```

[Back](#properties)
