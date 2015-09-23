# Reference Collection

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:B2";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	ctx.trackedObjects.add(range);
	range.load('address');
	return ctx.sync().then(function() {
		range.insert("Down");
		console.log(range.address); // Address should be updated to A3:B4
		return ctx.sync();
	});
});
```

### remove(rangeObject: range)

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:B2";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	ctx.trackedObjects.add(range);
	range.load('address');
	return ctx.sync().then(function() {
		range.insert("Down");
		console.log(range.address); // Address should be updated to A3:B4
		ctx.trackedObjects.remove(range); 
		return ctx.sync();
	});
});
```
