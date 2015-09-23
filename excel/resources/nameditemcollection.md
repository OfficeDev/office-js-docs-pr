# NamedItemCollection Object (JavaScript API for Excel)

_Applies to: Excel 2016, Office 2016_

A collection of all the nameditem objects that are part of the workbook.

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|items|[NamedItem[]](nameditem.md)|A collection of namedItem objects. Read-only.|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getItem(name: string)](#getitemname-string)|[NamedItem](nameditem.md)|Gets a nameditem object using its name|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details

### getItem(name: string)
Gets a nameditem object using its name

#### Syntax
```js
namedItemCollectionObject.getItem(name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|name|string|nameditem name.|

#### Returns
[NamedItem](nameditem.md)

#### Examples

```js
Excel.run(function (ctx) { 
	var nameditem = ctx.workbook.names.getItem(wSheetName);
	nameditem.load('type');
	return ctx.sync().then(function() {
			console.log(nameditem.type);
	});
});
```

```js
Excel.run(function (ctx) { 
	var nameditem = ctx.workbook.names.getItemAt(0);
	nameditem.load('name');
	return ctx.sync().then(function() {
			console.log(nameditem.name);
	});
});
```### load(param: object)
Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.

#### Syntax
```js
object.load(param);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|param|object|Optional. Accepts parameter and relationship names as delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns
void

#### Examples
```js

```
### Property access examples

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

