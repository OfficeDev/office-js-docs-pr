# ChartFill object (JavaScript API for Excel)

_Applies to: Excel 2016, Excel Online, Office 2016_

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

Clear the line format of the major gridlines on value axis of the chart named "Chart1".

```js
Excel.run(function (ctx) { 
	var gridlines = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").axes.valueaxis.majorGridlines;	
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
|color|string|HTML color code representing the color of the border line, of the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").|

#### Returns
void

#### Examples

Set the backGround color of Chart1 to red.

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
