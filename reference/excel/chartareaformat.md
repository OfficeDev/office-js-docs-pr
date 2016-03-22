# ChartAreaFormat object (JavaScript API for Excel)

Encapsulates the format properties for the overall chart area.

## Properties

None

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|fill|[ChartFill](chartfill.md)|Represents the fill format of an object, which includes background formatting information. Read-only.|
|font|[ChartFont](chartfont.md)|Represents the font attributes (font name, font size, color, etc.) for the current object. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in the JavaScript layer with property and object values specified in the parameter.|

## Method Details


### load(param: object)
Fills the proxy object created in the JavaScript layer with property and object values specified in the parameter.

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
