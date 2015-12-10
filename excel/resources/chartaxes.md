# ChartAxes object (JavaScript API for Excel)

_Applies to: Excel 2016, Excel Online, Office 2016_

Represents the chart axes.

## Properties

None

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|categoryAxis|[ChartAxis](chartaxis.md)|Represents the category axis in a chart. Read-only.|
|seriesAxis|[ChartAxis](chartaxis.md)|Represents the series axis of a 3-dimensional chart. Read-only.|
|valueAxis|[ChartAxis](chartaxis.md)|Represents the value axis in an axis. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in the JavaScript layer, with property and object values specified in the parameter.|

## Method Details

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
