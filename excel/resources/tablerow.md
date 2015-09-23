# Table Row
### delete()  

```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var row = ctx.workbook.tables.getItem(tableName).tableRows.getItemAt(2);
	row.delete();
	return ctx.sync(); 
	}); 
}); 
```

### getRange() 

```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var row = ctx.workbook.tables.getItem(tableName).tableRows.getItemAt(0);
	var rowRange = row.getRange();
	rowRange.load('address');
	return ctx.sync().then(function() {
		console.log(rowRange.address);
	});
}); 
```

### Getter and setter Table Row 

```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var row = ctx.workbook.tables.getItem(tableName).tableRows.getItem(0);
	row.load('index');
	return ctx.sync().then(function() {
		console.log(row.index);
	});
}); 
```

```js
Excel.run(function (ctx) { 
	var tables = ctx.workbook.tables;
	var newValues = [["New", "Values", "For", "New", "Row"]];
	var row = ctx.workbook.tables.getItem(tableName).tableRows.getItemAt(2);
	row.values = newValues;
	row.load('values');
	return ctx.sync().then(function() {
		console.log(row.values);
	});
}); 
```