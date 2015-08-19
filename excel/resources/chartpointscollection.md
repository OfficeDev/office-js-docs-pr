# ChartPointsCollection

A collection of all the chart points within a series inside a chart.

## [Properties](#getter-examples)
| Property	   | Type	|Description
|:---------------|:--------|:----------|
|count|int|Returns the number of chart points in the collection. Read-only.|
|items|[ChartPoint[]](chartpoint.md)|A collection of chartPoints objects. Read-only.|

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getItemAt(index: number)](#getitematindex-number)|[ChartPoint](chartpoint.md)|Retrieve a point based on its position within the series.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## API Specification

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
Set the border color for the first points in the points collection

```js
var ctx = new Excel.RequestContext();
var point = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").series.getItemAt(0).points;
points.getItemAt(0).format.fill.setSolidColor("8FBC8F");
ctx.executeAsync().then(function () {
	Console.log("Point Border Color Changed");
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

Get the names of points in the points collection
```js
var ctx = new Excel.RequestContext();
var pointsCollection = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").points;
pointsCollection.load(items);
ctx.executeAsync().then(function () {
	Console.log("Points Collection loaded");
});
```

Get the number of points

```js
var ctx = new Excel.RequestContext();
var pointsCollection = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").points;
pointsCollection.load(count);
ctx.executeAsync().then(function () {
	Console.log("points: Count= " + pointsCollection.count);
});

```

[Back](#properties)
