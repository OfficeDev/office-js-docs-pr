# ChartFont Object (JavaScript API for Excel)

This object represents the font attributes (font name, font size, color, etc.) for a chart object.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|bold|bool|Represents the bold status of font.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|color|string|HTML color code representation of the text color. E.g. #FF0000 represents Red.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|italic|bool|Represents the italic status of the font.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|Font name (e.g. "Calibri")|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|size|double|Size of the font (e.g. 11)|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|underline|string|Type of underline applied to the font. Possible values are: None, Single.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods
None


## Method Details

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
	title.format.font.underline = "None";
	return ctx.sync();
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

Set chart title to be Calbri, size 10, bold and in red. 

```js
Excel.run(function (ctx) { 
	var title = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").title;
	title.format.font.name = "Calibri";
	title.format.font.size = 12;
	title.format.font.color = "#FF0000";
	title.format.font.italic =  false;
	title.format.font.bold = true;
	title.format.font.underline = "None";
	return ctx.sync();
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
