# Binding Object (JavaScript API for Excel)

Represents an Office.js binding that is defined in the workbook.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|id|string|Represents binding identifier. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|type|string|Returns the type of the binding. Read-only. Possible values are: Range, Table, Text.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|Deletes the binding.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[getRange()](#getrange)|[Range](range.md)|Returns the range represented by the binding. Will throw an error if binding is not of the correct type.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getTable()](#gettable)|[Table](table.md)|Returns the table represented by the binding. Will throw an error if binding is not of the correct type.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getText()](#gettext)|string|Returns the text represented by the binding. Will throw an error if binding is not of the correct type.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### delete()
Deletes the binding.

#### Syntax
```js
bindingObject.delete();
```

#### Parameters
None

#### Returns
void

### getRange()
Returns the range represented by the binding. Will throw an error if binding is not of the correct type.

#### Syntax
```js
bindingObject.getRange();
```

#### Parameters
None

#### Returns
[Range](range.md)

#### Examples
Below example uses binding object to get the associated range.

```js
Excel.run(function (ctx) { 
	var binding = ctx.workbook.bindings.getItemAt(0);
	var range = binding.getRange();
	range.load('cellCount');
	return ctx.sync().then(function() {
		console.log(range.cellCount);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### getTable()
Returns the table represented by the binding. Will throw an error if binding is not of the correct type.

#### Syntax
```js
bindingObject.getTable();
```

#### Parameters
None

#### Returns
[Table](table.md)

#### Examples
```js
Excel.run(function (ctx) { 
	var binding = ctx.workbook.bindings.getItemAt(0);
	var table = binding.getTable();
	table.load('name');
	return ctx.sync().then(function() {
			console.log(table.name);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### getText()
Returns the text represented by the binding. Will throw an error if binding is not of the correct type.

#### Syntax
```js
bindingObject.getText();
```

#### Parameters
None

#### Returns
string

#### Examples

```js
Excel.run(function (ctx) { 
	var binding = ctx.workbook.bindings.getItemAt(0);
	var text = binding.getText();
	binding.load('text');
	return ctx.sync().then(function() {
		console.log(text);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### Property access examples

```js
Excel.run(function (ctx) { 
	var binding = ctx.workbook.bindings.getItemAt(0);
	binding.load('type');
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
