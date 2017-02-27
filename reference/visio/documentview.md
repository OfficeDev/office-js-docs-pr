# DocumentView object (JavaScript API for Visio)

Applies to: _Visio Online_

Represents the DocumentView class.

## Properties

| Property	   | Type	|Description|
|:---------------|:--------|:----------|
|disableHyperlinks|bool|Disable Hyperlinks.|
|disablePan|bool|Disable Pan.|
|disableZoom|bool|Disable Zoom.|
|hideDiagramBoundry|bool|Hide Diagram Boundry.|

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
