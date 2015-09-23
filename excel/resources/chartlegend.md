# Chart Legend
### Getter and setter

Get the `position` of Chart Legend from Chart1

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	var legend = chart.legend;
	legend.load('position');
	return ctx.sync().then(function() {
			console.log(legend.position);
	});
});
```

Set to show legend of Chart1 and make it on top of the chart.

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	chart.legend.visible = true;
	chart.legend.position = "top"; 
	chart.legend.overlay = false; 
	return ctx.sync().then(function() {
			console.log("Legend Shown ");
	});
});
``` 
