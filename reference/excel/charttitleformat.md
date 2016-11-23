# ChartTitleFormat object (JavaScript API for Excel)

Provides access to the office art formatting for chart title.

## Properties

None

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|fill|[ChartFill](chartfill.md)|Represents the fill format of an object, which includes background formating information. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|font|[ChartFont](chartfont.md)|Represents the font attributes (font name, font size, color, etc.) for an object. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

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
