# Range

### clear(applyTo: string)

Below example clears format and contents of the range. 

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "D:F";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	range.clear();
	return ctx.sync(); 
}); 
```

### delete(shift: string)

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "D:F";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	range.delete();
	return ctx.sync(); 
}); 
```

### getCell(row: number, column: number)

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:F8";
	var worksheet = ctx.workbook.worksheets.getItem(sheetName);
	var range = worksheet.getRange(rangeAddress);
	var cell = range.cell(0,0);
	cell.load('address');
	return ctx.sync().then(function() {
		console.log(cell.address);
	});
});
```

### getEntireColumn()

Note: the grid properties of the Range (values, numberFormat, formulas) contains `null` since the Range in question is unbounded.

```js

Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "D:F";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	var rangeEC = range.getEntireColumn();
	rangeEC.load('address');
	return ctx.sync().then(function() {
		console.log(rangeEC.address);
	});
});
```
### getBoundingRect(anotherRange: object)

```js

Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "D4:G6";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	var range = range.getBoundingRect("G4:H8");
	range.load('address');
	return ctx.sync().then(function() {
		console.log(range.address); // Prints Sheet1!D4:H8
	});
});
```

### getEntireRow()
```js

Excel.run(function (ctx) {
	var sheetName = "Sheet1";
	var rangeAddress = "D:F"; 
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	var rangeER = range.getEntireRow();
	rangeER.load('address');
	return ctx.sync().then(function() {
		console.log(rangeER.address);
	});
});
```
The grid properties of the Range (values, numberFormat, formulas) contains `null` since the Range in question is unbounded.

### getColumn(column: number)

```js

Excel.run(function (ctx) { 
	var sheetName = "Sheet19";
	var rangeAddress = "A1:F8";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getColumn(1);
	range.load('address');
	return ctx.sync().then(function() {
		console.log(range.address); // prints Sheet1!B1:B8
	});
});
```

### getLastCell()

```js

Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:F8";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getLastCell();
	range.load('address');
	return ctx.sync().then(function() {
		console.log(range.address); // prints Sheet1!F8
	});
});
```

### getLastColumn()

```js

Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:F8";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getLastColumn();
	range.load('address');
	return ctx.sync().then(function() {
		console.log(range.address); // prints Sheet1!F1:F8
	});
});
```

### getLastRow()

```js

Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:F8";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getLastRow();
	range.load('address');
	return ctx.sync().then(function() {
		console.log(range.address); // prints Sheet1!A8:F8
	});
});
```


### getRow(row: number)

```js

Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:F8";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getRow(1);
	range.load('address');
	return ctx.sync().then(function() {
		console.log(range.address); // prints Sheet1!A2:F2
	});
});
```

### getIntersection(anotherRange: object)

```js

Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:F8";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getIntersection("D4:G6");
	range.load('address');
	return ctx.sync().then(function() {
		console.log(range.address); // prints Sheet1!D4:F6
	});
});
```

### getOffsetRange(rowOffset: number, columnOffset: number)

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "D4:F6";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getOffsetRange(-1,4);
	range.load('address');
	return ctx.sync().then(function() {
		console.log(range.address); // prints Sheet1!H3:K5
	});
});
```

### getUsedRange()

```js

Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "D:F";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	var rangeUR = range.getUsedRange();
	rangeUR.load('address');
	return ctx.sync().then(function() {
		console.log(rangeUR.address);
	});
});
```

### insert(shift: string)

```js
	
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "F5:F10";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	range.insert();
	return ctx.sync(); 
	}); 
});
```

### select()

```js

Excel.run(function (ctx) {
	var sheetName = "Sheet1";
	var rangeAddress = "F5:F10"; 
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	range.select();
	return ctx.sync(); 
	}); 
});
```

### Getter and setter 

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
	var rangeName = 'MyRange';
	var range = ctx.workbook.names.getItem(rangeName).range;
	range.load('cellCount');
	return ctx.sync().then(function() {
		console.log(range.cellCount);
	});
});
```

The example below sets number-format, values and formulas on a grid that contains 2x3 grid.

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "F5:G7";
	var numberFormat = [[null, "d-mmm"], [null, "d-mmm"], [null, null]]
	var values = [["Today", 42147], ["Tomorrow", "5/24"], ["Difference in days", null]];
	var formulas = [[null,null], [null,null], [null,"=G6-G5"]];
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	range.numberFormat = numberFormat;
	range.values = values;
	range.formulas= formulas;
	range.load('text');
	return ctx.sync().then(function() {
		console.log(range.text);
	});
});
```
Get the worksheet containing the range. 

```js
Excel.run(function (ctx) { 
	var names = ctx.workbook.names;
	var namedItem = names.getItem('MyRange');
	range = namedItem.range;
	var rangeWorksheet = range.worksheet;
	rangeWorksheet.load('name');
	return ctx.sync().then(function() {
			console.log(rangeWorksheet.name);
	});
});
```
