# RangeFont Object (JavaScript API for Excel)

This object represents the font attributes (font name, font size, color, etc.) for an object.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|bold|bool|Represents the bold status of font.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|color|string|HTML color code representation of the text color. E.g. #FF0000 represents Red.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|italic|bool|Represents the italic status of the font.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|Font name (e.g. "Calibri")|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|size|double|Font size.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|underline|string|Type of underline applied to the font. Possible values are: None, Single, Double, SingleAccountant, DoubleAccountant.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods
None


## Method Details

### Property access examples

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "F:G";
	var worksheet = ctx.workbook.worksheets.getItem(sheetName);
	var range = worksheet.getRange(rangeAddress);
	var rangeFont = range.format.font;
	rangeFont.load('name');
	return ctx.sync().then(function() {
		console.log(rangeFont.name);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
The example below sets font name. 

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "F:G";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	range.format.font.name = 'Times New Roman';
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```