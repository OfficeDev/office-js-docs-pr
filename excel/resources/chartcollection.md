# ChartCollection

A collection of all the chart objects on a worksheet.

## [Properties](#getter-examples)
| Property	   | Type	|Description
|:---------------|:--------|:----------|
|count|int|Returns the number of charts in the worksheet. Read-only.|
|items|[Chart[]](chart.md)|A collection of chart objects. Read-only.|

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[add(type: string, sourceData: Range or string, seriesBy: string)](#addtype-string-sourcedata-range-or-string-seriesby-string)|[Chart](chart.md)|Creates a new chart.|
|[getItem(name: string)](#getitemname-string)|[Chart](chart.md)|Gets a chart using its name. If there are multiple charts with the same name, the first one will be returned.|
|[getItemAt(index: number)](#getitematindex-number)|[Chart](chart.md)|Gets a chart based on its position in the collection.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## API Specification

### add(type: string, sourceData: Range or string, seriesBy: string)
Creates a new chart.

#### Syntax
```js
chartCollectionObject.add(type, sourceData, seriesBy);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|type|string|Represents the type of a chart.  Possible values are: ColumnClustered, ColumnStacked, ColumnStacked100, BarClustered, BarStacked, BarStacked100, LineStacked, LineStacked100, LineMarkers, LineMarkersStacked, LineMarkersStacked100, PieOfPie, etc.|
|sourceData|Range or string|The address or name of the range that contains the source data. If an address or a worksheet-scoped name is used, it must include the worksheet name (e.g. "Sheet1!A5:B9"). |
|seriesBy|string|Optional. Specifies the way columns or rows are used as data series on the chart.  Possible values are: Auto, Columns, Rows|

#### Returns
[Chart](chart.md)

#### Examples

Add a chart of `chartType` "ColumnClustered" on worksheet "Charts" with `sourceData` from Range "A1:B4" and `seriresBy` is set to be "auto".

```js
var sheetName = "Sheet1";
var sourceData = sheetName + "!" + "A1:B4";
var ctx = new Excel.RequestContext();
var chart = ctx.workbook.worksheets.getItem(sheetName).charts.add("ColumnClustered", sourceData, "auto");
ctx.executeAsync().then(function () {
		Console.log("New Chart Added");
});
```


[Back](#methods)

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
var ctx = new Excel.RequestContext();
var chartname = 'Chart1';
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem(chartname);
ctx.executeAsync().then(function () {
		Console.log(chart.height);
});
```


```js
var ctx = new Excel.RequestContext();
var chartId = 'SamplChartId';
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem(chartId);
ctx.executeAsync().then(function () {
		Console.log(chart.height);
});
```



```js
var ctx = new Excel.RequestContext();
var lastPosition = ctx.workbook.worksheets.getItem("Sheet1").charts.count - 1;
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItemAt(lastPosition);
ctx.executeAsync().then(function () {
		Console.log(chart.name);
});
```


[Back](#methods)

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
var ctx = new Excel.RequestContext();
var lastPosition = ctx.workbook.worksheets.getItem("Sheet1").charts.count - 1;
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItemAt(lastPosition);
ctx.executeAsync().then(function () {
		Console.log(chart.name);
});
```


[Back](#methods)

### load(param: object)
Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.

#### Syntax
```js
object.load(param);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|param|object|Optional. Accepts parameter and relationship names as delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns
void

#### Examples
```js

```

[Back](#methods)

### Getter Examples

```js
var ctx = new Excel.RequestContext();
var charts = ctx.workbook.worksheets.getItem("Sheet1").charts;
charts.load(items);
ctx.executeAsync().then(function () {
	for (var i = 0; i < charts.items.length; i++)
	{
		Console.log(charts.items[i].name);
		Console.log(charts.items[i].index);
	}
});
```

Get the number of charts

```js
var ctx = new Excel.RequestContext();
var charts = ctx.workbook.worksheets.getItem("Sheet1").charts;
charts.load(count);
ctx.executeAsync().then(function () {
	Console.log("charts: Count= " + charts.count);
});
```


[Back](#properties)
