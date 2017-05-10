# Workbook Object (JavaScript API for Excel)

Workbook is the top level object which contains related workbook objects such as worksheets, tables, ranges, etc.

## Properties

None

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|application|[Application](application.md)|Represents Excel application instance that contains this workbook. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|bindings|[BindingCollection](bindingcollection.md)|Represents a collection of bindings that are part of the workbook. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|customXmlParts|[CustomXmlPartCollection](customxmlpartcollection.md)|Represents the collection of custom XML parts contained by this workbook. Read-only.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|functions|[Functions](functions.md)|Represents Excel application instance that contains this workbook. Read-only.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|names|[NamedItemCollection](nameditemcollection.md)|Represents a collection of workbook scoped named items (named ranges and constants). Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|pivotTables|[PivotTableCollection](pivottablecollection.md)|Represents a collection of PivotTables associated with the workbook. Read-only.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|settings|[SettingCollection](settingcollection.md)|Represents a collection of Settings associated with the workbook. Read-only.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|tables|[TableCollection](tablecollection.md)|Represents a collection of tables associated with the workbook. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|test|[Test](test.md)|For internal use only. Read-only.|[1.6](../requirement-sets/excel-api-requirement-sets.md)|
|worksheets|[WorksheetCollection](worksheetcollection.md)|Represents a collection of worksheets associated with the workbook. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getSelectedRange()](#getselectedrange)|[Range](range.md)|Gets the currently selected range from the workbook.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

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