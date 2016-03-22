# Workbook object (JavaScript API for Excel)

Workbook is the top-level object that contains related workbook objects such as worksheets, tables, ranges, etc.

## Properties

None

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|application|[Application](application.md)|Represents an instance of Excel application that contains this workbook. Read-only.|
|bindings|[BindingCollection](bindingcollection.md)|Represents a collection of bindings that are part of the workbook. Read-only.|
|functions|[Functions](functions.md)|Represents Excel application instance that contains this workbook. Read-only.|
|names|[NamedItemCollection](nameditemcollection.md)|Represents a collection of workbook scoped named items (named ranges and constants). Read-only.|
|tables|[TableCollection](tablecollection.md)|Represents a collection of tables associated with the workbook. Read-only.|
|worksheets|[WorksheetCollection](worksheetcollection.md)|Represents a collection of worksheets associated with the workbook. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getSelectedRange()](#getselectedrange)|[Range](range.md)|Gets the currently selected range from the workbook.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in the JavaScript layer, with property and object values specified in the parameter.|

## Method Details


### getSelectedRange()
Gets the currently selected range from the workbook.

#### Syntax
```js
workbookObject.getSelectedRange();
```

#### Parameters
None

#### Returns
[Range](range.md)

#### Examples

```js
Excel.run(function (ctx) { 
	var selectedRange = ctx.workbook.getSelectedRange();
	selectedRange.load('address');
	return ctx.sync().then(function() {
			console.log(selectedRange.address);
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
|param|object|Optional. Accepts parameter and relationship names as a delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns
void
