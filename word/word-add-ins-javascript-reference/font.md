# Font Object (JavaScript API for Word)

_Word 2016, Word for iPad, Word for Mac_

Represents a font.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|bold|bool|Gets or sets a value that indicates whether the font is bold. True if the font is formatted as bold, otherwise, false.|1.1||
|color|string|Gets or sets the color for the specified font. You can provide the value in the '#RRGGBB' format or the color name.|1.1||
|doubleStrikeThrough|bool|Gets or sets a value that indicates whether the font has a double strike through. True if the font is formatted as double strikethrough text, otherwise, false.|WordApiDesktop, 1.3||
|highlightColor|string|Gets or sets the highlight color for the specified font. You can provide the value as either in the '#RRGGBB' format or the color name.|1.1||
|italic|bool|Gets or sets a value that indicates whether the font is italicized. True if the font is italicized, otherwise, false.|1.1||
|name|string|Gets or sets a value that represents the name of the font.|1.1||
|strikeThrough|bool|Gets or sets a value that indicates whether the font has a strike through. True if the font is formatted as strikethrough text, otherwise, false.|1.1||
|subscript|bool|Gets or sets a value that indicates whether the font is a subscript. True if the font is formatted as subscript, otherwise, false.|1.1||
|superscript|bool|Gets or sets a value that indicates whether the font is a superscript. True if the font is formatted as superscript, otherwise, false.|1.1||

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|size|[float](float.md)|Gets or sets a value that represents the font size in points.|1.1||
|underline|[UnderlineType](underlinetype.md)|Gets or sets a value that indicates the font's underline type. 'None' if the font is not underlined.|1.1||

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.1|

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
