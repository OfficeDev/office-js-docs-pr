# BoundingBox object (JavaScript API for Visio)

Applies to: _Visio Online_
>**Note:** The Visio JavaScript APIs are currently in preview and are subject to change. The Visio JavaScript APIs are not currently supported for use in production environments.

Represents the BoundingBox of the shape.

## Properties

| Property	   | Type	|Description| Feedback|
|:---------------|:--------|:----------|:---|
|height|int|The distance between the top and bottom edges of the bounding box of the shape, excluding any data graphics associated with the shape.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-boundingBox-height)|
|width|int|The distance between the left and right edges of the bounding box of the shape, excluding any data graphics associated with the shape.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-boundingBox-width)|
|x|int|An integer that specifies the x-coordinate of the bounding box.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-boundingBox-x)|
|y|int|An integer that specifies the y-coordinate of the bounding box.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-boundingBox-y)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Feedback|
|:---------------|:--------|:----------|:---|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-boundingBox-load)|

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
