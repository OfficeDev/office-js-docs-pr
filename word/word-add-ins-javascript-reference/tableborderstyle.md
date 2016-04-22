# TableBorderStyle Object (JavaScript API for Word)

_Word 2016, Word for iPad, Word for Mac_

Specifies the border style

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|color|string|Gets or sets the table border color, as a hex value or name.|
|type|string|Gets or sets the type of the table border style. Possible values are: Mixed, None, Single, Thick, Double, Hairline, Dotted, Dashed, DotDashed, Dot2Dashed, Triple, ThinThickSmall, ThickThinSmall, ThinThickThinSmall, ThinThickMed, ThickThinMed, ThinThickThinMed, ThinThickLarge, ThickThinLarge, ThinThickThinLarge, Wave, DoubleWave, DashedSmall, DashDotStroked, ThreeDEmboss, ThreeDEngrave.|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|width|[float](float.md)|Gets or sets the width, in points, of the table border style.|

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
|:---------------|:--------|:----------|
|param|object|Optional. Accepts parameter and relationship names as delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns
void
