# ChartLineFormat

Enapsulates the formatting options for line elements.

## [Properties](#setter-examples)
| Property	   | Type	|Description
|:---------------|:--------|:----------|
|color|string|HTML color code representing the color of lines in the chart.|

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[clear()](#clear)|void|Clear the line format of a chart element.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## API Specification

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

Clear the line format of the major gridlines on value axis of the Chart named "Chart1"

```js
var ctx = new Excel.RequestContext();
var gridlines = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").axes.valueaxis.majorGridlines;	

gridlines.format.line.clear();
ctx.executeAsync().then(function () {
		Console.log"Chart Major Gridlines Format Cleared");
});
```

[Back](#methods)

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

### Setter Examples

Set chart major gridlines on value axis to be red.
```js
var ctx = new Excel.RequestContext();
var gridlines = ctx.workbook.worksheets.getItem("Sheet1").charts.axes.valueaxis.majorGridlines;


[Back](#properties)
