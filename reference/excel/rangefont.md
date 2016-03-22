# RangeFont object (JavaScript API for Excel)

This object represents the font attributes (font name, font size, color, etc.) for an object.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|bold|bool|Represents the bold setting for font.|
|color|string|HTML color code representation of the text color, for example, #FF0000 represents red.|
|italic|bool|Represents the italic setting for font.|
|name|string|Font name, for example, "Calibri".|
|size|double|Font size.|
|underline|string|Type of underline applied to the font. Possible values are: None, Single, Double, SingleAccountant, DoubleAccountant.|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in the JavaScript layer, with property and object values specified in the parameter.|

## Method Details


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
The example below sets the font name. 

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
