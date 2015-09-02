# ChartFill

Represents the fill formatting for a chart element.

## Properties
None

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[clear()](#clear)|void|Clear the fill color of a chart element.|
|[setSolidColor(color: string)](#setsolidcolorcolor-string)|void|Sets the fill formatting of a chart element to a uniform color.|

## API Specification

### clear()
Clear the fill color of a chart element.

#### Syntax
```js
chartFillObject.clear();
```

#### Parameters
None

#### Returns
void

#### Examples

Clear the line format of the major Gridlines on value axis of the Chart named "Chart1"

```js
var ctx = new Excel.RequestContext();
var gridlines = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").axes.valueaxis.majorGridlines;	

gridlines.format.line.clear();
ctx.executeAsync().then(function () {
		Console.log"Chart Major Gridlines Format Cleared");
});
```

[Back](#methods)

### setSolidColor(color: string)
Sets the fill formatting of a chart element to a uniform color.

#### Syntax
```js
chartFillObject.setSolidColor(color);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|color|string|HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|

#### Returns
void

#### Examples

Set BackGround Color of Chart1 to be red.
```js
var ctx = new Excel.RequestContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	

chart.format.fill.setSolidColor("#FF0000");

ctx.executeAsync().then(function () {
		Console.log("Chart1 Background Color Changed.");
});
```

[Back](#methods)

