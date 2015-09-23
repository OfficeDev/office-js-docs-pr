# NamedItem Collection
A collection of all the nameditem objects that are part of the workbook. 

### Getter 

```js
Excel.run(function (ctx) { 
	var nameditems = ctx.workbook.names;
	nameditems.load('items');
	return ctx.sync().then(function() {
		for (var i = 0; i < nameditems.items.length; i++)
		{
			console.log(nameditems.items[i].name);
			console.log(nameditems.items[i].index);
		}
	});
});
```

Get the number of nameditems.

```js
Excel.run(function (ctx) { 
	var nameditems = ctx.workbook.names;
	nameditems.load('count');
	return ctx.sync().then(function() {
		console.log("nameditems: Count= " + nameditems.count);
	});
});

```

### getItem(name: string)

```js
Excel.run(function (ctx) { 
	var nameditem = ctx.workbook.names.getItem(wSheetName);
	nameditem.load('type');
	return ctx.sync().then(function() {
			console.log(nameditem.type);
	});
});
```
### getItemAt(index: number)

```js
Excel.run(function (ctx) { 
	var nameditem = ctx.workbook.names.getItemAt(0);
	nameditem.load('name');
	return ctx.sync().then(function() {
			console.log(nameditem.name);
	});
});
```