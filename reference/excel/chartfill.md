# ChartFill Object (JavaScript API for Excel)

Represents the fill formatting for a chart element.

## Properties

None

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[clear()](#clear)|void|Clear the fill color of a chart element.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[setSolidColor(color: string)](#setsolidcolorcolor-string)|void|Sets the fill formatting of a chart element to a uniform color.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


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
Excel.run(function (ctx) { 
	var gridlines = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").axes.valueAxis.majorGridlines;	
	gridlines.format.line.clear();
	return ctx.sync().then(function() {
			console.log("Chart Major Gridlines Format Cleared");
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

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
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	

	chart.format.fill.setSolidColor("#FF0000");

	return ctx.sync().then(function() {
			console.log("Chart1 Background Color Changed.");
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
