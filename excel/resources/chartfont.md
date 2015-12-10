# ChartFont object (JavaScript API for Excel)

_Applies to: Excel 2016, Excel Online, Office 2016_

This object represents the font attributes (font name, font size, color, etc.) for a chart object.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|bold|bool|Represents the bold setting for font.|
|color|string|HTML color code representation of the text color, e.g., #FF0000 represents red.|
|italic|bool|Represents the italic setting for font.|
|name|string|Font name, e.g., Calibri.|
|size|double|Size of the font, e.g., 11.|
|underline|string|Type of underline applied to the font. Possible values are: None, Single.|

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

Use chart title as an example.

```js
Excel.run(function (ctx) { 
	var title = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").title;
	title.format.font.name = "Calibri";
	title.format.font.size = 12;
	title.format.font.color = "#FF0000";
	title.format.font.italic =  false;
	title.format.font.bold = true;
	title.format.font.underline = false;
	return ctx.sync();
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

Set the chart title format to be Calibri, size 10, bold and red. 

```js
Excel.run(function (ctx) { 
	var title = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").title;
	title.format.font.name = "Calibri";
	title.format.font.size = 12;
	title.format.font.color = "#FF0000";
	title.format.font.italic =  false;
	title.format.font.bold = true;
	title.format.font.underline = false;
	return ctx.sync();
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
