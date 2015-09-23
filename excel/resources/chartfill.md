# Chart Fill Format
### setSolidColor(color: string)

Set BackGround Color of Chart1 to be red.

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	

	chart.format.fill.setSolidColor("#FF0000");

	return ctx.sync().then(function() {
			console.log("Chart1 Background Color Changed.");
	});
});
```
### clear()

Clear the line format of the major Gridlines on value axis of the Chart named "Chart1"

```js
Excel.run(function (ctx) { 
	var gridlines = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").axes.valueaxis.majorGridlines;	
	gridlines.format.line.clear();
	return ctx.sync().then(function() {
			console.log"Chart Major Gridlines Format Cleared");
	});
});

```
