# Workbook

Workbook is the top level object which contains related workbook objects such as worksheets, tables, ranges, etc.

## Properties
None

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|application|[Application](application.md)|Represents Excel application instance that contains this workbook. Read-only.|
|bindings|[BindingCollection](bindingcollection.md)|Represents a collection of bindings that are part of the workbook. Read-only.|
|names|[NamedItemCollection](nameditemcollection.md)|Represents a collection of workbook scoped named items (named ranges and constants). Read-only.|
|tables|[TableCollection](tablecollection.md)|Represents a collection of tables associated with the workbook. Read-only.|
|worksheets|[WorksheetCollection](worksheetcollection.md)|Represents a collection of worksheets associated with the workbook. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getSelectedRange()](#getselectedrange)|[Range](range.md)|Gets the currently selected range from the workbook.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## API Specification

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
var ctx = new Excel.RequestContext();
var selectedRange = ctx.workbook.getSelectedRange();
ctx.executeAsync().then(function () {
		Console.log(selectedRange.address);
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

