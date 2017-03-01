# Selection object (JavaScript API for Visio)

Applies to: _Visio Online_

Represents the Selection in the page.

## Properties

None

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|shapes|[ShapeCollection](shapecollection.md)|Gets the Shapes of the Selection Read-only.|

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
