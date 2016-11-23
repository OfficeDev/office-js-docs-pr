# NamedItemCollection object (JavaScript API for Excel)

A collection of all the nameditem objects that are part of the workbook.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[NamedItem[]](nameditem.md)|A collection of namedItem objects. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getItem(name: string)](#getitemname-string)|[NamedItem](nameditem.md)|Gets a nameditem object using its name|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNull(name: string)](#getitemornullname-string)|[NamedItem](nameditem.md)|Gets a nameditem object using its name. If the nameditem object does not exist, the returned object's isNull property will be true.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### getItem(name: string)
Gets a nameditem object using its name

#### Syntax
```js
namedItemCollectionObject.getItem(name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|name|string|nameditem name.|

#### Returns
[NamedItem](nameditem.md)

#### Examples

```js
Excel.run(function (ctx) { 
	var sheetName = 'Sheet1';
	var nameditem = ctx.workbook.names.getItem(sheetName);
	nameditem.load('type');
	return ctx.sync().then(function() {
			console.log(nameditem.type);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
### getItemOrNull(name: string)
Gets a nameditem object using its name. If the nameditem object does not exist, the returned object's isNull property will be true.

#### Syntax
```js
namedItemCollectionObject.getItemOrNull(name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|name|string|nameditem name.|

#### Returns
[NamedItem](nameditem.md)

### load(param: object)
Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.

#### Syntax
```js
object.load(param);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|param|object|Optional. Accepts parameter and relationship names as delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns
void
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
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


