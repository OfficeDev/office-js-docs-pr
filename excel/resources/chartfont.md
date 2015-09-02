# ChartFont

This object represents the font attributes (font name, font size, color, etc.) for a chart object.

## [Properties](#setter-examples)
| Property	   | Type	|Description
|:---------------|:--------|:----------|
|bold|bool|Represents the bold status of font.|
|color|string|HTML color code representation of the text color. E.g. #FF0000 represents Red.|
|italic|bool|Represents the italic status of the font.|
|name|string|Font name (e.g. "Calibri")|
|size|double|Size of the font (e.g. 11)|
|underline|string|Type of underline applied to the font. Possible values are: None, Single.|

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## API Specification

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

#### Examples
```js

```

[Back](#methods)

### Setter Examples

Use chart title as an example.

```js
chartObject.title.format.font.name = "Calibri";
chartObject.title.format.font.size = 12;

[Back](#properties)
