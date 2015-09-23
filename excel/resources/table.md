# Table

### delete()
```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var table = ctx.workbook.tables.getItem(tableName);
	table.delete();
	return ctx.sync(); 
}); 
```

### getDataBodyRange()
```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var table = ctx.workbook.tables.getItem(tableName);
	var tableDataRange = table.getDataBodyRange();
	tableDataRange.load('address')
	return ctx.sync().then(function() {
			console.log(tableDataRange.address);
	});
});
```
### getHeaderRowRange()
```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var table = ctx.workbook.tables.getItem(tableName);
	var tableHeaderRange = table.getHeaderRowRange();
	tableHeaderRange.load('address');
	return ctx.sync().then(function() {
		console.log(tableHeaderRange.address);
	});
});
```

### getRange()
```js
Excel.run(function (ctx) { 
	var table = ctx.workbook.tables.getItem(tableName);
	var tableRange = table.getRange();
	tableRange.load('address');	
	return ctx.sync().then(function() {
			console.log(tableRange.address);
	});
});
```

### getTotalRowRange()
```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var table = ctx.workbook.tables.getItem(tableName);
	var tableTotalsRange = table.getTotalRowRange();
	tableTotalsRange.load('address');	
	return ctx.sync().then(function() {
			console.log(tableTotalsRange.address);
	});
});
```

### getter and setter

Get a table by name. 

```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var table = ctx.workbook.tables.getItem(tableName);
	table.load('index')
	return ctx.sync().then(function() {
			console.log(table.index);
	});
});
```

Get a table by index.

```js
Excel.run(function (ctx) { 
	var index = 0;
	var table = ctx.workbook.tables.getItemAt(0);
	table.name('name')
	return ctx.sync().then(function() {
			console.log(table.name);
	});
});
```

Set table style. 

```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var table = ctx.workbook.tables.getItem(tableName);
	table.name = 'Table1-Renamed';
	table.showTotals = false;
	table.tableStyle = 'TableStyleMedium2';
	table.load('tableStyle');
	return ctx.sync().then(function() {
			console.log(table.tableStyle);
	});
});
```