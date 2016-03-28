# RangeFormat object (JavaScript API for Excel)

A format object encapsulating the range's font, fill, borders, alignment, and other properties.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|columnWidth|double|Gets or sets the width of all colums within the range. If the column widths are not uniform, null will be returned.|
|horizontalAlignment|string|Represents the horizontal alignment for the specified object. Possible values are: General, Left, Center, Right, Fill, Justify, CenterAcrossSelection, Distributed.|
|rowHeight|double|Gets or sets the height of all rows in the range. If the row heights are not uniform null will be returned.|
|verticalAlignment|string|Represents the vertical alignment for the specified object. Possible values are: Top, Center, Bottom, Justify, Distributed.|
|wrapText|bool|Indicates Excel text control is set to wrap text in the object. A null value indicates the entire range doesn’t use a uniform wrap text setting.|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|borders|[RangeBorderCollection](rangebordercollection.md)|Collection of border objects that apply to the overall range selected Read-only.|
|fill|[RangeFill](rangefill.md)|Returns the fill object defined on the overall range. Read-only.|
|font|[RangeFont](rangefont.md)|Returns the font object defined on the overall range selected Read-only.|
|protection|[FormatProtection](formatprotection.md)|Returns the format protection object for a range. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[autofitColumns()](#autofitcolumns)|void|Changes the width of the columns of the current range to achieve the best fit, based on the current data in the columns.|
|[autofitRows()](#autofitrows)|void|Changes the height of the rows of the current range to achieve the best fit, based on the current data in the columns.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in the JavaScript layer, with property and object values specified in the parameter.|

## Method Details


### autofitColumns()
Changes the width of the columns of the current range to achieve the best fit, based on the current data in the columns.

#### Syntax
```js
rangeFormatObject.autofitColumns();
```

#### Parameters
None

#### Returns
void

### autofitRows()
Changes the height of the rows of the current range to achieve the best fit, based on the current data in the columns.

#### Syntax
```js
rangeFormatObject.autofitRows();
```

#### Parameters
None

#### Returns
void

### load(param: object)
Fills the proxy object created in the JavaScript layer, with property and object values specified in the parameter.

#### Syntax
```js
object.load(param);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|param|object|Optional. Accepts parameter and relationship names as a delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns
void
### Property access examples

This example prints all of the format properties of a range. 

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "F:G";
	var worksheet = ctx.workbook.worksheets.getItem(sheetName);
	var range = worksheet.getRange(rangeAddress);
	range.load(["format/*", "format/fill", "format/borders", "format/font"]);
	return ctx.sync().then(function() {
		console.log(range.format.wrapText);
		console.log(range.format.fill.color);
		console.log(range.format.font.name);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

The example below sets font name and fill color of a range and wraps text. 

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "F:G";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	range.format.wrapText = true;
	range.format.font.name = 'Times New Roman';
	range.format.fill.color = '0000FF';
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

The example below adds a grid border around the range.

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "F:G";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	range.format.borders('InsideHorizontal').lineStyle = 'Continuous';
	range.format.borders('InsideVertical').lineStyle = 'Continuous';
	range.format.borders('EdgeBottom').lineStyle = 'Continuous';
	range.format.borders('EdgeLeft').lineStyle = 'Continuous';
	range.format.borders('EdgeRight').lineStyle = 'Continuous';
	range.format.borders('EdgeTop').lineStyle = 'Continuous';
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
