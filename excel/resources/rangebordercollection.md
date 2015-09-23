# Range Border Collection 

### getItem(index: string)
```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:F8";
	var worksheet = ctx.workbook.worksheets.getItem(sheetName);
	var range = worksheet.getRange(rangeAddress);
	var borderName = 'EdgeTop';
	var border = range.format.borders.getItem(borderName);
	border.load('style');
	return ctx.sync().then(function() {
			console.log(border.style);
	});
});

```

### getItemAt(index: number)
```js

Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:F8";
	var worksheet = ctx.workbook.worksheets.getItem(sheetName);
	var range = worksheet.getRange(rangeAddress);
	var border = ctx.workbook.borders.getItemAt(0);
	border.load('sideIndex');
	return ctx.sync().then(function() {
			console.log(border.sideIndex);
	});
});

```

### Getter and setter

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:F8";
	var worksheet = ctx.workbook.worksheets.getItem(sheetName);
	var range = worksheet.getRange(rangeAddress);
	var borders = range.format.borders;
	border.load('items');
	return ctx.sync().then(function() {
		console.log(borders.count);
		for (var i = 0; i < borders.items.length; i++)
		{
			console.log(borders.items[i].sideIndex);
		}
	});
});

```
The example below adds grid border around the range.

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:F8";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	range.format.borders.getItem('InsideHorizontal').style = 'Continuous';
	range.format.borders.getItem('InsideVertical').style = 'Continuous';
	range.format.borders.getItem('EdgeBottom').style = 'Continuous';
	range.format.borders.getItem('EdgeLeft').style = 'Continuous';
	range.format.borders.getItem('EdgeRight').style = 'Continuous';
	range.format.borders.getItem('EdgeTop').style = 'Continuous';
	return ctx.sync(); 
}); 