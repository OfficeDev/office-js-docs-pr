# Chart object (JavaScript API for Excel)

Represents a chart object in a workbook.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|height|double|Represents the height, in points, of the chart object.|
|id|string|Gets a chart based on its position in the collection. Read-only.|
|left|double|The distance, in points, from the left side of the chart to the worksheet origin.|
|name|string|Represents the name of a chart object.|
|top|double|Represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).|
|width|double|Represents the width, in points, of the chart object.|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|axes|[ChartAxes](chartaxes.md)|Represents chart axes. Read-only.|
|dataLabels|[ChartDataLabels](chartdatalabels.md)|Represents the data labels on the chart. Read-only.|
|format|[ChartAreaFormat](chartareaformat.md)|Encapsulates the format properties for the chart area. Read-only.|
|legend|[ChartLegend](chartlegend.md)|Represents the legend for the chart. Read-only.|
|series|[ChartSeriesCollection](chartseriescollection.md)|Represents either a single series or a series collection in the chart. Read-only.|
|title|[ChartTitle](charttitle.md)|Represents the title of the specified chart, including the text, visibility, position, and formatting of the title. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[delete()](#delete)|void|Deletes the chart object.|
|[getImage(height: number, width: number, fittingMode: string)](#getimageheight-number-width-number-fittingmode-string)|[System.IO.Stream](system.io.stream.md)|Renders the chart as a base64-encoded image by scaling the chart to fit the specified dimensions.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in the JavaScript layer with property and object values specified in the parameter.|
|[setData(sourceData: Range, seriesBy: string)](#setdatasourcedata-range-seriesby-string)|void|Resets the source data for the chart.|
|[setPosition(startCell: Range or string, endCell: Range or string)](#setpositionstartcell-range-or-string-endcell-range-or-string)|void|Positions the chart relative to cells on the worksheet.|

## Method Details


### delete()
Deletes the chart object.

#### Syntax
```js
chartObject.delete();
```

#### Parameters
None

#### Returns
void

#### Examples
```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	chart.delete();
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getImage(height: number, width: number, fittingMode: string)
Renders the chart as a base64-encoded image by scaling the chart to fit the specified dimensions.

#### Syntax
```js
chartObject.getImage(height, width, fittingMode);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|height|number|Optional. (Optional) The desired height of the resulting image.|
|width|number|Optional. (Optional) The desired width of the resulting image.|
|fittingMode|string|Optional. (Optional) The method used to scale the chart to the specified to the specified dimensions (if both height and width are set)."  Possible values are: Fit, FitAndCenter, Fill|

#### Returns
[System.IO.Stream](system.io.stream.md)

#### Examples
```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	var image = chart.getImage();
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```





### load(param: object)
Fills the proxy object created in the JavaScript layer with property and object values specified in the parameter.

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

### setData(sourceData: Range, seriesBy: string)
Resets the source data for the chart.

#### Syntax
```js
chartObject.setData(sourceData, seriesBy);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|sourceData|Range|The Range object corresponding to the source data.|
|seriesBy|string|Optional. Specifies the way columns or rows are used as data series on the chart. Can be one of the following: Auto (default), Rows, Columns.  Possible values are: Auto, Columns, Rows.|

#### Returns
void

#### Examples

Set the `sourceData` to be "A1:B4" and `seriesBy` to be "Columns".

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	var sourceData = "A1:B4";
	chart.setData(sourceData, "Columns");
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### setPosition(startCell: Range or string, endCell: Range or string)
Positions the chart relative to cells on the worksheet.

#### Syntax
```js
chartObject.setPosition(startCell, endCell);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|startCell|Range or string|The start cell. This is where the chart will be moved to. The start cell is the top-left or top-right cell, depending on the user's left-to-right display settings.|
|endCell|Range or string|Optional. The end cell. If specified, the chart's width and height is set to fully cover up this cell/range.|

#### Returns
void

#### Examples


```js
Excel.run(function (ctx) { 
	var sheetName = "Charts";
	var sourceData = sheetName + "!" + "A1:B4";
	var chart = ctx.workbook.worksheets.getItem(sheetName).charts.add("pie", sourceData, "auto");
	chart.width = 500;
	chart.height = 300;
	chart.setPosition("C2", null);
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### Property access examples

Get a chart named "Chart1".

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	chart.load('name');
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

Update a chart, including renaming, positioning, and resizing.

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	chart.name="New Name";
	chart.top = 100;
	chart.left = 100;
	chart.height = 200;
	chart.weight = 200;
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

Give the chart a new name; Resize the chart to 200 points in both height and weight. Move Chart1 to 100 points to the top and left. 

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");
	chart.name="New Name";	
	chart.top = 100;
	chart.left = 100;
	chart.height =200;
	chart.width =200;
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

