# Chart Series Collection

### Getter ChartSeries Collection
Getting the names of series in the series collection.

```js
Excel.run(function (ctx) { 
	var seriesCollection = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").series;
	seriesCollection.load('items');
	return ctx.sync().then(function() {
		for (var i = 0; i < seriesCollection.items.length; i++)
		{
			console.log(seriesCollection.items[i].name);
		}
	});
});
```

Get the number of chart series in collection.

```js
Excel.run(function (ctx) { 
	var seriesCollection = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").series;
	seriesCollection.load('count');
	return ctx.sync().then(function() {
		console.log("series: Count= " + seriesCollection.count);
	});
});

```

### getItemAt(index: number)

Get the name of the first series in the series collection.

```js
Excel.run(function (ctx) { 
	var seriesCollection = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").series;
	seriesCollection.load('items');
	return ctx.sync().then(function() {
		console.log(seriesCollection.items[0].name);
	});
});
```

