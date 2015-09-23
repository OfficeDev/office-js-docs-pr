# Chart Data Labels
### Getter

Make Series Name shown in Datalabels and set the `position` of datalabels to be "top";

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	chart.datalabels.visible = true;
	chart.datalabels.position = "top";
	chart.datalabels.ShowSeriesName = true;
	return ctx.sync().then(function() {
			console.log("Datalabels Shown");
	});
});

```
