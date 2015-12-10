# ChartPoint object (JavaScript API for Excel)

_Applies to: Excel 2016, Excel Online, Office 2016_

Represents a point of a series in a chart.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|value|object|Returns the value of a chart point. Read-only.|

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|format|[ChartPointFormat](chartpointformat.md)|Encapsulates the format properties chart point. Read-only.|

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
