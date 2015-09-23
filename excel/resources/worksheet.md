# Worksheet

### activate()

```js
Excel.run(function (ctx) { 
	var wSheetName = 'Sheet1';
	var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
	worksheet.activate();
	return ctx.sync(); 
	}); 
}); 
```

### delete()

```js
Excel.run(function (ctx) { 
	var wSheetName = 'Sheet1';
	var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
	worksheet.delete();
	return ctx.sync(); 
	}); 
}); 
```

### getCell(row: number, column: number)

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:F8";
	var worksheet = ctx.workbook.worksheets.getItem(sheetName);
	var cell = worksheet.getCell(0,0);
	cell.load('address');
	return ctx.sync().then(function() {
		console.log(cell.address);
	});
});
```

### getRange(address: string)
Below example uses range address to get the range object.

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:F8";
	var worksheet = ctx.workbook.worksheets.getItem(sheetName);
	var range = worksheet.getRange(rangeAddress);
	range.load('cellCount');
	return ctx.sync().then(function() {
		console.log(range.cellCount);
	});
});
```

Below example uses a named-range to get the range object.

```js

Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeName = 'MyRange';
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeName);
	range.load('address');
	return ctx.sync().then(function() {
		console.log(range.address);
	});
});
```
### getUsedRange()

```js
Excel.run(function (ctx) { 
	var wSheetName = 'Sheet1';
	var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
	var usedRange = worksheet.getUsedRange();
	usedRange.load('address');
	return ctx.sync().then(function() {
			console.log(usedRange.address);
	});
});
```

### Getter and setter Worksheet

Get worksheet properties based on sheet name.

```js
Excel.run(function (ctx) { 
	var wSheetName = 'Sheet1';
	var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
	worksheet.load('position')
	return ctx.sync().then(function() {
			console.log(worksheet.position);
	});
});
```

Set worksheet position. 

```js
Excel.run(function (ctx) { 
	var wSheetName = 'Sheet1';
	var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
	worksheet.position = 0;
	return ctx.sync(); 
}); 
```

