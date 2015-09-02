# NamedItem

Represents a defined name for a range of cells or value. Names can be primitive named objects (as seen in the type below), range object, reference to a range. This object can be used to obtain range object associated with names.

## [Properties](#getter-examples)
| Property	   | Type	|Description
|:---------------|:--------|:----------|
|name|string|The name of the object. Read-only.|
|type|string|Indicates what type of reference is associated with the name. Read-only. Possible values are: String, Integer, Double, Boolean, Range.|
|value|object|Represents the formula that the name is defined to refer to. E.g. =Sheet14!$B$2:$H$12, =4.75, etc. Read-only.|
|visible|bool|Specifies whether the object is visible or not.|

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getRange()](#getrange)|[Range](range.md)|Returns the range object that is associated with the name. Throws an exception if the named item's type is not a range.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## API Specification

### getRange()
Returns the range object that is associated with the name. Throws an exception if the named item's type is not a range.

#### Syntax
```js
namedItemObject.getRange();
```

#### Parameters
None

#### Returns
[Range](range.md)

#### Examples

Returns the Range object that is associated with the name. `null` if the name is not of the type `Range`. Note: This API currently supports only the Workbook scoped items.**

```js
var ctx = new Excel.RequestContext();
var names = ctx.workbook.names;
var range = names.getItem('MyRange').getRange();
range.load(address);
ctx.executeAsync().then(function () {
		Console.log(range.address);
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
var names = ctx.workbook.names;
var namedItem = names.getItem('MyRange');
namedItem.load(type);
ctx.executeAsync().then(function () {
		Console.log(namedItem.type);
});
```

[Back](#properties)
