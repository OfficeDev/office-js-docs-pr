# ConditionalRangeFont Object (JavaScript API for Excel)

This object represents the font attributes (font style,, color, etc.) for an object.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|bold|bool|Represents the bold status of font.|[1.6](../requirement-sets/excel-api-requirement-sets.md)|
|color|string|HTML color code representation of the text color. E.g. #FF0000 represents Red.|[1.6](../requirement-sets/excel-api-requirement-sets.md)|
|italic|bool|Represents the italic status of the font.|[1.6](../requirement-sets/excel-api-requirement-sets.md)|
|strikethrough|bool|Represents the strikethrough status of the font.|[1.6](../requirement-sets/excel-api-requirement-sets.md)|
|underline|string|Type of underline applied to the font. Possible values are: None, Single, Double.|[1.6](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[clear()](#clear)|void|Resets the font formats.|[1.6](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### clear()
Resets the font formats.

#### Syntax
```js
conditionalRangeFontObject.clear();
```

#### Parameters
None

#### Returns
void
