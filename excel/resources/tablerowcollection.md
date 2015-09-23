# Table Row Collection

### add(index: number, values: object[][])

```js
Excel.run(function (ctx) { 
	var tables = ctx.workbook.tables;
	var values = [["Sample", "Values", "For", "New", "Row"]];
	var row = tables.getItem("Table1").rows.add(null, values);
	row.load('index');
	return ctx.sync().then(function() {
		console.log(row.index);
	});
}); 
```
### getItemAt(index: number)

```js
Excel.run(function (ctx) { 
	var tablerow = ctx.workbook.tables.getItem('Table1').rows.getItemAt(0);
	tablerow.load('name');
	return ctx.sync().then(function() {
			console.log(tablerow.name);
	});
}); 
```
### Getter tablerow Collection

```js
Excel.run(function (ctx) { 
	var tablerows = ctx.workbook.tables.getItem('Table1').rows;
	tablerows.load('items');
	return ctx.sync().then(function() {
		console.log("tablerows Count: " + tablerows.count);
		for (var i = 0; i < tablerows.items.length; i++)
		{
			console.log(tablerows.items[i].index);
		}
	});
}); 
```