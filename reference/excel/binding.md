# Binding object (JavaScript API for Excel)

Represents an Office.js binding that is defined in the workbook.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|id|string|Represents binding identifier. Read-only.|
|type|string|Returns the type of the binding. Read-only. Possible values are: Range, Table, Text.|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getRange()](#getrange)|[Range](range.md)|Returns the range represented by the binding. Will throw an error if binding is not of the correct type.|
|[getTable()](#gettable)|[Table](table.md)|Returns the table represented by the binding. Will throw an error if binding is not of the correct type.|
|[getText()](#gettext)|string|Returns the text represented by the binding. Will throw an error if binding is not of the correct type.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in the JavaScript layer, with property and object values specified in the parameter.|

## Method Details


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
The example below uses a binding object to get the associated range.

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
	ctx.load('text');
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


### load(param: object)
Fills the proxy object created in the JavaScript layer, with property and object values specified in the parameter.

#### Syntax
```js
object.load(param);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|param|object|Optional. Accepts parameter and relationship names as a delimited string or an array. Or, accepts a [loadOption](loadoption.md) object.|

#### Returns
void
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
