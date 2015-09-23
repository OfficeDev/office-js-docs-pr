# Table Collection

### add(address: string, hasHeaders: bool)

```js
Excel.run(function (ctx) { 
	var table = ctx.workbook.tables.add('Sheet1!A1:E7', true);
	table.load('name');
	return ctx.sync().then(function() {
		console.log(table.name);
	});
});

```
### getItem(id: object)

```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var table = ctx.workbook.tables.getItem(tableName);
	return ctx.sync().then(function() {
			console.log(table.index);
	});
});
```

### getItemAt(index: number)

```js
Excel.run(function (ctx) { 
	var table = ctx.workbook.tables.getItemAt(0);
	return ctx.sync().then(function() {
			console.log(table.name);
	});
});
```

### Getter 

```js
Excel.run(function (ctx) { 
	var tables = ctx.workbook.tables;
	tables.load('items');
	return ctx.sync().then(function() {
		console.log("tables Count: " + tables.count);
		for (var i = 0; i < tables.items.length; i++)
		{
			console.log(tables.items[i].name);
		}
	});
});
```

Get the number of tables

```js
Excel.run(function (ctx) { 
	var tables = ctx.workbook.tables;
	tables.load('count');
	return ctx.sync().then(function() {
		console.log(tables.count);
	});
});

```