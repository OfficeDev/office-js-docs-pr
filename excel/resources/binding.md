# Binding

Represents an Office.js binding that is defined in the workbook.

## [Properties](#getter-examples)
| Property	   | Type	|Description
|:---------------|:--------|:----------|
|id|string|Represents binding identifier. Read-only.|
|type|string|Returns the type of the binding. Read-only. Possible values are: Range, Table, Text.|

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getRange()](#getrange)|[Range](range.md)|Returns the range represented by the binding. Will throw an error if binding is not of the correct type.|
|[getTable()](#gettable)|[Table](table.md)|Returns the table represented by the binding. Will throw an error if binding is not of the correct type.|
|[getText()](#gettext)|string|Returns the text represented by the binding. Will throw an error if binding is not of the correct type.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## API Specification

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
var ctx = new Excel.RequestContext();
var binding = ctx.workbook.bindings.getItemAt(0);
var range = binding.getRange();
range.load(cellCount);
ctx.executeAsync().then(function() {
	Console.log(range.cellCount);
});
```


[Back](#methods)

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
var ctx = new Excel.RequestContext();

var binding = ctx.workbook.bindings.getItemAt(0);
var table = binding.getTable();
table.load(name);
ctx.executeAsync().then(function () {
		Console.log(table.name);
});
```


[Back](#methods)

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
var ctx = new Excel.RequestContext();
var binding = ctx.workbook.bindings.getItemAt(0);
var text = binding.getText();
ctx.load(text);
ctx.executeAsync().then(function() {
	Console.log(text);
});
```


[Back](#methods)

### load(param: object)
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

[Back](#methods)

### Getter Examples

```js
var ctx = new Excel.RequestContext();
var binding = ctx.workbook.bindings.getItemAt(0);
binding.load(type);
ctx.executeAsync().then(function() {
	Console.log(binding.type);
});
```
[Back](#properties)
