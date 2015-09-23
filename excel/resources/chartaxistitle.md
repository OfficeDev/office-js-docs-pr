# AxisTitle

### Getter and setter
Get the `text` of Chart Axis Title from the value axis of Chart1.

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	var title = chart.axes.valueaxis.title;
	title.load('text');
	return ctx.sync().then(function() {
			console.log(title.text);
	});
});
```

Add "Values" as the title for the value Axis

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	chart.axes.valueaxis.title.text = "Values";
	return ctx.sync().then(function() {
			console.log("Axis Title Added ");
	});
});
```
