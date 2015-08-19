# RangeFormat

A format object encapsulating the range's font, fill, borders, alignment, and other properties.

## [Properties](#getter-and-setter-examples)
| Property	   | Type	|Description
|:---------------|:--------|:----------|
|horizontalAlignment|string|Represents the horizontal alignment for the specified object. Possible values are: General, Left, Center, Right, Fill, Justify, CenterAcrossSelection, Distributed.|
|verticalAlignment|string|Represents the vertical alignment for the specified object. Possible values are: Top, Center, Bottom, Justify, Distributed.|
|wrapText|bool|Indicates if Excel wraps the text in the object. A null value indicates that the entire range doesn't have uniform wrap setting|

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|borders|[RangeBorderCollection](rangebordercollection.md)|Collection of border objects that apply to the overall range selected Read-only.|
|fill|[RangeFill](rangefill.md)|Returns the fill object defined on the overall range. Read-only.|
|font|[RangeFont](rangefont.md)|Returns the font object defined on the overall range selected Read-only.|

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

### Getter and Setter Examples

Below example selects all of the Range's format properties. 

```js
var sheetName = "Sheet1";
var rangeAddress = "F:G";
var ctx = new Excel.RequestContext();
var worksheet = ctx.workbook.worksheets.getItem(sheetName);
var range = worksheet.getRange(rangeAddress);
range.load(format, format/fill, format/borders, format/font);
ctx.executeAsync().then(function() {
	Console.log(range.format.wrapText);
	Console.log(range.format.fill.color);
	Console.log(range.format.font.name);

});
```

The example below sets font name, fill color and wraps text. 

```js
var sheetName = "Sheet1";
var rangeAddress = "F:G";
var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
range.format.wrapText = true;
range.format.font.name = 'Times New Roman';
range.format.fill.color = '0000FF';
ctx.executeAsync();
```

The example below adds grid border around the range.

```js
var sheetName = "Sheet1";
var rangeAddress = "F:G";
var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
range.format.borders('InsideHorizontal').lineStyle = 'Continuous';
range.format.borders('InsideVertical').lineStyle = 'Continuous';
range.format.borders('EdgeBottom').lineStyle = 'Continuous';
range.format.borders('EdgeLeft').lineStyle = 'Continuous';
range.format.borders('EdgeRight').lineStyle = 'Continuous';
range.format.borders('EdgeTop').lineStyle = 'Continuous';
ctx.executeAsync();
```

[Back](#properties)
