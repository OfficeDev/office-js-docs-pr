# InlinePictureCollection object (JavaScript API for Word)

Contains a collection of [inlinePicture](inlinepicture.md) objects.

_Applies to: Word 2016_

## Properties
| Property	   | Type	|Description
|:---------------|:--------|:----------|
|items|[InlinePicture[]](inlinepicture.md)|A collection of inlinePicture objects. Read-only.|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method details

### load(param: object)
Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.

#### Syntax
```js
object.load(param);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|param|object|Optional. Accepts parameter and relationship names as delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns
void


