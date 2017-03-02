# NamedItem Object (JavaScript API for Excel)

Represents a defined name for a range of cells or value. Names can be primitive named objects (as seen in the type below), range object, reference to a range. This object can be used to obtain range object associated with names.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|comment|string|Represents the comment associated with this name.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|The name of the object. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|scope|string|Indicates whether the name is scoped to the workbook or to a specific worksheet. Read-only. Possible values are: Equal, Greater, GreaterEqual, Less, LessEqual, NotEqual.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|type|string|Indicates the type of the value returned by the name's formula. Read-only. Possible values are: String, Integer, Double, Boolean, Range.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|value|object|Represents the value computed by the name's formula. For a named range, will return the range address. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|visible|bool|Specifies whether the object is visible or not.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|worksheet|[Worksheet](worksheet.md)|Returns the worksheet on which the named item is scoped to. Throws an error if the items is scoped to the workbook instead. Read-only.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|worksheetOrNullObject|[Worksheet](worksheet.md)|Returns the worksheet on which the named item is scoped to. Returns a null object if the item is scoped to the workbook instead. Read-only.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|Deletes the given name.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getRange()](#getrange)|[Range](range.md)|Returns the range object that is associated with the name. Throws an error if the named item's type is not a range.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getRangeOrNullObject()](#getrangeornullobject)|[Range](range.md)|Returns the range object that is associated with the name. Returns a null object if the named item's type is not a range.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### delete()
Deletes the given name.

#### Syntax
```js
namedItemObject.delete();
```

#### Parameters
None

#### Returns
void

### getRange()
Returns the range object that is associated with the name. Throws an error if the named item's type is not a range.

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
Excel.run(function (ctx) { 
	var names = ctx.workbook.names;
	var range = names.getItem('MyRange').getRange();
	range.load('address');
	return ctx.sync().then(function() {
			console.log(range.address);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### getRangeOrNullObject()
Returns the range object that is associated with the name. Returns a null object if the named item's type is not a range.

#### Syntax
```js
namedItemObject.getRangeOrNullObject();
```

#### Parameters
None

#### Returns
[Range](range.md)
### Property access examples

```js
Excel.run(function (ctx) { 
	var names = ctx.workbook.names;
	var namedItem = names.getItem('MyRange');
	namedItem.load('type');
	return ctx.sync().then(function() {
			console.log(namedItem.type);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
