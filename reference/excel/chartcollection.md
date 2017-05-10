# ChartCollection Object (JavaScript API for Excel)

A collection of all the chart objects on a worksheet.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|count|int|Returns the number of charts in the worksheet. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|items|[Chart[]](chart.md)|A collection of chart objects. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[add(type: string, sourceData: object, seriesBy: string)](#addtype-string-sourcedata-object-seriesby-string)|[Chart](chart.md)|Creates a new chart.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount()](#getcount)|int|Returns the number of charts in the worksheet.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(name: string)](#getitemname-string)|[Chart](chart.md)|Gets a chart using its name. If there are multiple charts with the same name, the first one will be returned.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[Chart](chart.md)|Gets a chart based on its position in the collection.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNullObject(name: string)](#getitemornullobjectname-string)|[Chart](chart.md)|Gets a chart using its name. If there are multiple charts with the same name, the first one will be returned.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### add(type: string, sourceData: object, seriesBy: string)
Creates a new chart.

#### Syntax
```js
chartCollectionObject.add(type, sourceData, seriesBy);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|type|string|Represents the type of a chart.  Possible values are: ColumnClustered, ColumnStacked, ColumnStacked100, BarClustered, BarStacked, BarStacked100, LineStacked, LineStacked100, LineMarkers, LineMarkersStacked, LineMarkersStacked100, PieOfPie, etc.|
|sourceData|object|The Range object corresponding to the source data.|
|seriesBy|string|Optional. Specifies the way columns or rows are used as data series on the chart.  Possible values are: Auto, Columns, Rows|

#### Returns
[Chart](chart.md)

#### Examples

Add a chart of `chartType` "ColumnClustered" on worksheet "Charts" with `sourceData` from Range "A1:B4" and `seriresBy` is set to be "auto".

```js
Excel.run(function (ctx) { 
	var rangeSelection = "A1:B4";
	var range = ctx.workbook.worksheets.getItem(sheetName)
		.getRange(rangeSelection);
	var chart = ctx.workbook.worksheets.getItem(sheetName)
		.charts.add("ColumnClustered", range, "auto");	return ctx.sync().then(function() {
			console.log("New Chart Added");
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### getCount()
Returns the number of charts in the worksheet.

#### Syntax
```js
chartCollectionObject.getCount();
```

#### Parameters
None

#### Returns
int

### getItem(name: string)
Gets a chart using its name. If there are multiple charts with the same name, the first one will be returned.

#### Syntax
```js
chartCollectionObject.getItem(name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|name|string|Name of the chart to be retrieved.|

#### Returns
[Chart](chart.md)

#### Examples

```js
Excel.run(function (ctx) { 
	var chartname = 'Chart1';
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem(chartname);
	return ctx.sync().then(function() {
			console.log(chart.height);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


#### Examples

```js
Excel.run(function (ctx) { 
	var chartId = 'SamplChartId';
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem(chartId);
	return ctx.sync().then(function() {
			console.log(chart.height);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```



#### Examples

```js
Excel.run(function (ctx) { 
	var lastPosition = ctx.workbook.worksheets.getItem("Sheet1").charts.count - 1;
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItemAt(lastPosition);
	return ctx.sync().then(function() {
			console.log(chart.name);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### getItemAt(index: number)
Gets a chart based on its position in the collection.

#### Syntax
```js
chartCollectionObject.getItemAt(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|Index value of the object to be retrieved. Zero-indexed.|

#### Returns
[Chart](chart.md)

#### Examples

```js
Excel.run(function (ctx) { 
	var lastPosition = ctx.workbook.worksheets.getItem("Sheet1").charts.count - 1;
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItemAt(lastPosition);
	return ctx.sync().then(function() {
			console.log(chart.name);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### getItemOrNullObject(name: string)
Gets a chart using its name. If there are multiple charts with the same name, the first one will be returned.

#### Syntax
```js
chartCollectionObject.getItemOrNullObject(name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|name|string|Name of the chart to be retrieved.|

#### Returns
[Chart](chart.md)
### Property access examples

```js
Excel.run(function (ctx) { 
	var charts = ctx.workbook.worksheets.getItem("Sheet1").charts;
	charts.load('items');
	return ctx.sync().then(function() {
		for (var i = 0; i < charts.items.length; i++)
		{
			console.log(charts.items[i].name);
		}
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

Get the number of charts

```js
Excel.run(function (ctx) { 
	var charts = ctx.workbook.worksheets.getItem("Sheet1").charts;
	charts.load('count');
	return ctx.sync().then(function() {
		console.log("charts: Count= " + charts.count);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

