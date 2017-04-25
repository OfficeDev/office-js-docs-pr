# NamedItemCollection Object (JavaScript API for Excel)

A collection of all the nameditem objects that are part of the workbook or worksheet, depending on how it was reached.

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
|[add(name: string, reference: Range or string, comment: string)](#addname-string-reference-range-or-string-comment-string)|[NamedItem](nameditem.md)|Adds a new name to the collection of the given scope.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[addFormulaLocal(name: string, formula: string, comment: string)](#addformulalocalname-string-formula-string-comment-string)|[NamedItem](nameditem.md)|Adds a new name to the collection of the given scope using the user's locale for the formula.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount()](#getcount)|int|Gets the number of named items in the collection.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(name: string)](#getitemname-string)|[NamedItem](nameditem.md)|Gets a nameditem object using its name|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNullObject(name: string)](#getitemornullobjectname-string)|[NamedItem](nameditem.md)|Gets a nameditem object using its name. If the nameditem object does not exist, will return a null object.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### add(name: string, reference: Range or string, comment: string)
Adds a new name to the collection of the given scope.

#### Syntax
```js
namedItemCollectionObject.add(name, reference, comment);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|name|string|The name of the named item.|
|reference|Range or string|The formula or the range that the name will refer to.|
|comment|string|Optional. The comment associated with the named item|

#### Returns
[NamedItem](nameditem.md)

### addFormulaLocal(name: string, formula: string, comment: string)
Adds a new name to the collection of the given scope using the user's locale for the formula.

#### Syntax
```js
namedItemCollectionObject.addFormulaLocal(name, formula, comment);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|name|string|The "name" of the named item.|
|formula|string|The formula in the user's locale that the name will refer to.|
|comment|string|Optional. The comment associated with the named item|

#### Returns
[NamedItem](nameditem.md)

### getCount()
Gets the number of named items in the collection.

#### Syntax
```js
namedItemCollectionObject.getCount();
```

#### Parameters
None

#### Returns
int

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
### getItemOrNullObject(name: string)
Gets a nameditem object using its name. If the nameditem object does not exist, will return a null object.

#### Syntax
```js
namedItemCollectionObject.getItemOrNullObject(name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|name|string|nameditem name.|

#### Returns
[NamedItem](nameditem.md)
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


