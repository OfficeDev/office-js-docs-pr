# ChartPointsCollection object (JavaScript API for Excel)

_Applies to: Excel 2016, Excel Online, Office 2016_

A collection of all the chart points within a series inside a chart.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|count|int|Returns the number of chart points in the collection. Read-only.|
|items|[ChartPoint[]](chartpoint.md)|A collection of chartPoints objects. Read-only.|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getItemAt(index: number)](#getitematindex-number)|[ChartPoint](chartpoint.md)|Retrieve a point based on its position within the series.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in the JavaScript layer, with property and object values specified in the parameter.|

## Method Details

### getItemAt(index: number)
Retrieve a point based on its position within the series.

#### Syntax
```js
chartPointsCollectionObject.getItemAt(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|Index value of the object to be retrieved. Zero-indexed.|

#### Returns
[ChartPoint](chartpoint.md)

#### Examples
Set the border color for the first points in the points collection.

```js
Excel.run(function (ctx) { 
	var point = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").series.getItemAt(0).points;
	points.getItemAt(0).format.fill.setSolidColor("#8FBC8F");
	return ctx.sync().then(function() {
		console.log("Point Border Color Changed");
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### load(param: object)
Fills the proxy object created in the JavaScript layer, with property and object values specified in the parameter.

#### Syntax
```js
object.load(param);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|param|object|Optional. Accepts parameter and relationship names as a delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns
void
### Property access examples

Get the names of points in the points collection.

```js
Excel.run(function (ctx) { 
	var pointsCollection = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").points;
	pointsCollection.load('items');
	return ctx.sync().then(function() {
		console.log("Points Collection loaded");
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

Get the number of points.

```js
Excel.run(function (ctx) { 
	var pointsCollection = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").points;
	pointsCollection.load('count');
	return ctx.sync().then(function() {
		console.log("points: Count= " + pointsCollection.count);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
