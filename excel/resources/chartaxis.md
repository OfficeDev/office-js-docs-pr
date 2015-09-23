# Chart Axis

### Getter and setter
Get the `maximum` of Chart Axis from Chart1

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	var axis = chart.axes.valueaxis;
	axis.load('maximum');
	return ctx.sync().then(function() {
			console.log(axis.maximum);
	});
});
```

Set the  `maximum`,  `minimum`,  `majorunit`, `minorunit` of valueaxis. 

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	chart.axes.valueaxis.maximum = 5;
	chart.axes.valueaxis.minimum = 0;
	chart.axes.valueaxis.majorunit = 1;
	chart.axes.valueaxis.minorunit = 0.2;
	return ctx.sync().then(function() {
			console.log("Axis Settings Changed");
	});
});
```
