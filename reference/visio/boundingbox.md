# BoundingBox object (JavaScript API for Visio)

Applies to: _Visio Online_

Represents the BoundingBox of the shape.

## Properties

| Property	   | Type	|Description|
|:---------------|:--------|:----------|
|height|int|The distance between the top and bottom edges of the bounding box of the shape, excluding any data graphics associated with the shape.|
|width|int|The distance between the left and right edges of the bounding box of the shape, excluding any data graphics associated with the shape.|
|x|int|An integer that specifies the x-coordinate of the bounding box.|
|y|int|An integer that specifies the y-coordinate of the bounding box.|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

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
