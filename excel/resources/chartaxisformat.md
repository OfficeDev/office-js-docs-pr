# ChartAxisFormat object (JavaScript API for Excel)

_Applies to: Excel 2016, Excel Online, Office 2016_

Encapsulates the format properties for the chart axis.

## Properties

None

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|font|[ChartFont](chartfont.md)|Represents the font attributes such as font name, font size, color, etc., of the chart axis element. Read-only.|
|line|[ChartLineFormat](chartlineformat.md)|Represents chart line formatting. Read-only.|

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
