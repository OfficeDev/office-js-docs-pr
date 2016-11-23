# ChartLegendFormat object (JavaScript API for Excel)

Encapsulates the format properties of a chart legend.

## Properties

None

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|fill|[ChartFill](chartfill.md)|Represents the fill format of an object, which includes background formating information. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|font|[ChartFont](chartfont.md)|Represents the font attributes such as font name, font size, color, etc. of a chart legend. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

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
