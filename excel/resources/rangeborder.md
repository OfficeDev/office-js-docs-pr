# RangeBorder

Represents the border of an object.

## [Properties](#getter-and-setter-examples)
| Property	   | Type	|Description
|:---------------|:--------|:----------|
|color|string|HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|
|sideIndex|string|Constant value that indicates the specific side of the border. Read-only. Possible values are: EdgeTop, EdgeBottom, EdgeLeft, EdgeRight, InsideVertical, InsideHorizontal, DiagonalDown, DiagonalUp.|
|style|string|One of the constants of line style specifying the line style for the border. Possible values are: None, Continuous, Dash, DashDot, DashDotDot, Dot, Double, SlantDashDot.|
|weight|string|Specifies the weight of the border around a range. Possible values are: Hairline, Thin, Medium, Thick.|

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

### Getter and Setter Examples

```js
var sheetName = "Sheet1";
var rangeAddress = "A1:F8";
var ctx = new Excel.RequestContext();
var worksheet = ctx.workbook.worksheets.getItem(sheetName);
var range = worksheet.getRange(rangeAddress);
var borders = range.format.borders;
borders.load(items);
ctx.executeAsync().then(function () {
	Console.log(borders.count);
	for (var i = 0; i < borders.items.length; i++)
	{
		Console.log(borders.items[i].sideIndex);
	}
});
```
The example below adds grid border around the range.

```js
var sheetName = "Sheet1";
var rangeAddress = "A1:F8";
var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
range.format.borders.getItem('InsideHorizontal').style = 'Continuous';
range.format.borders.getItem('InsideVertical').style = 'Continuous';
range.format.borders.getItem('EdgeBottom').style = 'Continuous';
range.format.borders.getItem('EdgeLeft').style = 'Continuous';
range.format.borders.getItem('EdgeRight').style = 'Continuous';
range.format.borders.getItem('EdgeTop').style = 'Continuous';
ctx.executeAsync();
```



[Back](#properties)
