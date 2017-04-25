# BindingCollection Object (JavaScript API for Excel)

Represents the collection of all the binding objects that are part of the workbook.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|count|int|Returns the number of bindings in the collection. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|items|[Binding[]](binding.md)|A collection of binding objects. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[add(range: Range or string, bindingType: string, id: string)](#addrange-range-or-string-bindingtype-string-id-string)|[Binding](binding.md)|Add a new binding to a particular Range.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[addFromNamedItem(name: string, bindingType: string, id: string)](#addfromnameditemname-string-bindingtype-string-id-string)|[Binding](binding.md)|Add a new binding based on a named item in the workbook.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[addFromSelection(bindingType: string, id: string)](#addfromselectionbindingtype-string-id-string)|[Binding](binding.md)|Add a new binding based on the current selection.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount()](#getcount)|int|Gets the number of bindings in the collection.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(id: string)](#getitemid-string)|[Binding](binding.md)|Gets a binding object by ID.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[Binding](binding.md)|Gets a binding object based on its position in the items array.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNullObject(id: string)](#getitemornullobjectid-string)|[Binding](binding.md)|Gets a binding object by ID. If the binding object does not exist, will return a null object.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### add(range: Range or string, bindingType: string, id: string)
Add a new binding to a particular Range.

#### Syntax
```js
bindingCollectionObject.add(range, bindingType, id);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|range|Range or string|Range to bind the binding to. May be an Excel Range object, or a string. If string, must contain the full address, including the sheet name|
|bindingType|string|Type of binding.  Possible values are: Range, Table, Text|
|id|string|Name of binding.|

#### Returns
[Binding](binding.md)

### addFromNamedItem(name: string, bindingType: string, id: string)
Add a new binding based on a named item in the workbook.

#### Syntax
```js
bindingCollectionObject.addFromNamedItem(name, bindingType, id);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|name|string|Name from which to create binding.|
|bindingType|string|Type of binding.  Possible values are: Range, Table, Text|
|id|string|Name of binding.|

#### Returns
[Binding](binding.md)

### addFromSelection(bindingType: string, id: string)
Add a new binding based on the current selection.

#### Syntax
```js
bindingCollectionObject.addFromSelection(bindingType, id);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|bindingType|string|Type of binding.  Possible values are: Range, Table, Text|
|id|string|Name of binding.|

#### Returns
[Binding](binding.md)

### getCount()
Gets the number of bindings in the collection.

#### Syntax
```js
bindingCollectionObject.getCount();
```

#### Parameters
None

#### Returns
int

### getItem(id: string)
Gets a binding object by ID.

#### Syntax
```js
bindingCollectionObject.getItem(id);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|id|string|Id of the binding object to be retrieved.|

#### Returns
[Binding](binding.md)

#### Examples

Create a table binding to monitor data changes in the table. When data is changed, the background color of the table will be changed to orange.

```js
function addEventHandler() {
	//Create Table1
Excel.run(function (ctx) { 
	ctx.workbook.tables.add("Sheet1!A1:C4", true);
	return ctx.sync().then(function() {
			 console.log("My Diet Data Inserted!");
	})
	.catch(function (error) {
			 console.log(JSON.stringify(error));
	});
});
	//Create a new table binding for Table1
Office.context.document.bindings.addFromNamedItemAsync("Table1", Office.CoercionType.Table, { id: "myBinding" }, function (asyncResult) {
	if (asyncResult.status == "failed") {
		console.log("Action failed with error: " + asyncResult.error.message);
	}
	else {
		// If succeeded, then add event handler to the table binding.
		Office.select("bindings#myBinding").addHandlerAsync(Office.EventType.BindingDataChanged, onBindingDataChanged);
	}
});
}
	
// when data in the table is changed, this event will be triggered.
function onBindingDataChanged(eventArgs) {
Excel.run(function (ctx) { 
	// highlight the table in orange to indicate data has been changed.
	ctx.workbook.bindings.getItem(eventArgs.binding.id).getTable().getDataBodyRange().format.fill.color = "Orange";
	return ctx.sync().then(function() {
			console.log("The value in this table got changed!");
	})
	.catch(function (error) {
			console.log(JSON.stringify(error));
	});
});
}

```



#### Examples
```js
Excel.run(function (ctx) { 
	var lastPosition = ctx.workbook.bindings.count - 1;
	var binding = ctx.workbook.bindings.getItemAt(lastPosition);
	binding.load('type')
	return ctx.sync().then(function() {
			console.log(binding.type); 
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### getItemAt(index: number)
Gets a binding object based on its position in the items array.

#### Syntax
```js
bindingCollectionObject.getItemAt(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|Index value of the object to be retrieved. Zero-indexed.|

#### Returns
[Binding](binding.md)

#### Examples
```js
Excel.run(function (ctx) { 
	var lastPosition = ctx.workbook.bindings.count - 1;
	var binding = ctx.workbook.bindings.getItemAt(lastPosition);
	binding.load('type')
	return ctx.sync().then(function() {
			console.log(binding.type); 
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### getItemOrNullObject(id: string)
Gets a binding object by ID. If the binding object does not exist, will return a null object.

#### Syntax
```js
bindingCollectionObject.getItemOrNullObject(id);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|id|string|Id of the binding object to be retrieved.|

#### Returns
[Binding](binding.md)
### Property access examples

```js
Excel.run(function (ctx) { 
	var bindings = ctx.workbook.bindings;
	bindings.load('items');
	return ctx.sync().then(function() {
		for (var i = 0; i < bindings.items.length; i++)
		{
			console.log(bindings.items[i].id);
		}
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
Get the number of bindings

```js
Excel.run(function (ctx) { 
	var bindings = ctx.workbook.bindings;
	bindings.load('count');
	return ctx.sync().then(function() {
		console.log("Bindings: Count= " + bindings.count);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
