# ChartAxes object (JavaScript API for Excel)

Represents the chart axes.

## Properties

None

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|categoryAxis|[ChartAxis](chartaxis.md)|Represents the category axis in a chart. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|seriesAxis|[ChartAxis](chartaxis.md)|Represents the series axis of a 3-dimensional chart. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|valueAxis|[ChartAxis](chartaxis.md)|Represents the value axis in an axis. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### load(param: object)
Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.

#### Syntax
```js
object.load(param);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|param|object|Optional. Accepts parameter and relationship names as delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns
void
