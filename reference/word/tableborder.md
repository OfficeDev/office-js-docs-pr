# TableBorder Object (JavaScript API for Word)

_Word 2016, Word for iPad, Word for Mac, Word Online_

Specifies the border style

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|color|string|Gets or sets the table border color, as a hex value or name.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|[type](enums.md)|string|Gets or sets the type of the table border. Possible values are: Mixed, None, Single, Thick, Double, Hairline, Dotted, Dashed, DotDashed, Dot2Dashed, Triple, ThinThickSmall, ThickThinSmall, ThinThickThinSmall, ThinThickMed, ThickThinMed, ThinThickThinMed, ThinThickLarge, ThickThinLarge, ThinThickThinLarge, Wave, DoubleWave, DashedSmall, DashDotStroked, ThreeDEmboss, ThreeDEngrave.|[1.3](../requirement-sets/word-api-requirement-sets.md)|
|width|float|Gets or sets the width, in points, of the table border. Not applicable to table border types that have fixed widths.|[1.3](../requirement-sets/word-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[1.1](../requirement-sets/word-api-requirement-sets.md)|

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
