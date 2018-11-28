---
title: Work with workbooks using the Excel JavaScript API
description: ''
ms.date: 11/27/2018
---


# Work with workbooks using the Excel JavaScript API

This article provides code samples that show how to perform common tasks with workbooks using the Excel JavaScript API. For the complete list of properties and methods that the **Workbook** object supports, see [Workbook Object (JavaScript API for Excel)](/javascript/api/excel/excel.workbook). This article also covers workbook-level actions performed through the [Application](/javascript/api/excel/excel.application) object.

The Workbook object is the entry point for your add-in to interact with Excel. It maintains collections of worksheets, tables, PivotTables, and more, through which Excel data is accessed and changed. The [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) object gives your add-in access to all the workbook's data through indivual worksheets. Specifically, it lets your add-in add worksheets, navigate among them, and assign handlers to worksheet events. The article [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md) describes how to access and edit worksheets.

## Get the active cell or selected range

The Workbook object contains two methods that get a range of cells the user or add-in has selected: `getActiveCell()` and `getSelectedRange()`. `getActiveCell()` gets the active cell from the workbook as a [Range object](/javascript/api/excel/excel.range). The following example shows a call to `getActiveCell()`, followed by the cell's address being printed to the console.

```js
Excel.run(function (context) {
    var activeCell = context.workbook.getActiveCell();
	activeCell.load("address");

    return context.sync().then(function () {
		console.log("The active cell is " + activeCell.address);
	});
}).catch(errorHandlerFunction);
```

The `getSelectedRange()` method returns the currently selected single range. If multiple ranges are selected, an InvalidSelection error is thrown. The following example shows a call to `getSelectedRange()` that then sets the range's fill color to yellow.

```js
Excel.run(function(context) {
	var range = context.workbook.getSelectedRange();
	range.format.fill.color = "yellow";
	return context.sync();
}).catch(errorHandlerFunction);
```

## Create a workbook

Your add-in can create a new workbook, separate from the Excel instance in which the add-in is currently running. The Excel object has the `createWorkbook` method for this purpose. When this method is called, the new workbook is immediately opened and displayed in a new instance of Excel. Your add-in remains open and running with the previous workbook.

```js
Excel.createWorkbook();
```

The `createWorkbook` method can also create a copy of an existing workbook. The method accepts a base64-encoded string representation of an .xlsx file as an optional parameter. The resulting workbook will be a copy of that file, assuming the string argument is a valid .xlsx file.

You can get your add-in’s current workbook as a base64-encoded string by using [file slicing](/javascript/api/office/office.document#getfileasync-filetype--options--callback-). The [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) class can be used to convert a file into the required base64-encoded string, as demonstrated in the following example. 

```js
var myFile = document.getElementById("file");
var reader = new FileReader();

reader.onload = (function (event) {
	Excel.run(function (context) {
		// strip off the metadata before the base64-encoded string
		var startIndex = event.target.result.indexOf("base64,");
		var mybase64 = event.target.result.substr(startIndex + 7);

		Excel.createWorkbook(mybase64);
		return context.sync();
	}).catch(errorHandlerFunction);
});

// read in the file as a data URL so we can parse the base64-encoded string
reader.readAsDataURL(myFile.files[0]);
```

## Protect the workbook's structure

Your add-in can control a user's ability to edit the workbook's structure. The Workbook object's `protection` property is a [WorkbookProtection](/javascript/api/excel/excel.workbookprotection) object with a `protect()` method. The following example shows a basic scenario toggling the protection of the workbook's structure. 

```js
Excel.run(function (context) {
	var workbook = context.workbook;
	workbook.load("protection/protected");

	return context.sync().then(function() {
		if (!workbook.protection.protected) {
			workbook.protection.protect();
		}
	});
}).catch(errorHandlerFunction);
```

The `protect` method accepts an optional string parameter. This string represents the password needed for a user to bypass protection and change the workbook's structure.

Protection can also be set at the worksheet level to prevent unwanted data editing. For more information, see the **Data protection** section of the [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md#data-protection) article.

> [!NOTE] 
> For more information about workbook protection in Excel, see the [Protect a workbook](https://support.office.com/article/Protect-a-workbook-7E365A4D-3E89-4616-84CA-1931257C1517) article.

## Access document properties

Workbook objects have access to the Office file metadata, which is known as the [document properties](https://support.office.com/article/View-or-change-the-properties-for-an-Office-file-21D604C2-481E-4379-8E54-1DD4622C6B75). The Workbook object's `properties` property is a [DocumentProperties](/javascript/api/excel/excel.documentproperties) object containing these metadata values. The following example shows how to set the **author** property.

```js
Excel.run(function (context) {
	var docProperties = context.workbook.properties;
	docProperties.author = "Alex";
	return context.sync();
}).catch(errorHandlerFunction);
```

You can also define custom properties. The DocumentProperties object contains a `custom` property that represents a collection of key-value pairs for user-defined properties. The following example shows how to create a custom property named **Introduction** with the value "Hello", then retrieve it.

```js
Excel.run(function (context) {
	var customDocProperties = context.workbook.properties.custom;
	customDocProperties.add("Introduction", "Hello");
	return context.sync();
}).catch(errorHandlerFunction);

[...]

Excel.run(function (context) {
	var customDocProperties = context.workbook.properties.custom;
	var customProperty = customDocProperties.getItem("Introduction");
	customProperty.load("key, value");

	return context.sync().then(function() {
        console.log("Custom key  : " + customProperty.key); // "Introduction"
        console.log("Custom value : " + customProperty.value); // "Hello"
	});
}).catch(errorHandlerFunction);
```

## Access document settings

A workbook's settings are similar to the collection of custom properties. The difference is settings are unique to a single Excel file and add-in pairing, whereas properties are solely connected to the file. The following example shows how to create and access a setting.

```js
Excel.run(function (context) {
	var settings = context.workbook.settings;
	settings.add("NeedsReview", true);
	var needsReview = settings.getItem("NeedsReview");
	needsReview.load("value");

	return context.sync().then(function() {
		console.log("Workbook needs review : " + needsReview.value);
	});
}).catch(errorHandlerFunction);
```

## Control calculation behavior

### Set calculation mode

By default, Excel recalculates formula results whenever a referenced cell is changed. Your add-in's performance may benefit from adjusting this calculation behavior. The Application object has a `calculationMode` property of type `CalculationMode`. It can be set to the following values:

 - `automatic`: The default recalculation behavior where Excel calculates new formula results every time the relevant data is changed.
 - `automaticExceptTables`: Same as `automatic`, except any changes made to values in tables are ignored.
 - `manual`: Calculations only occur when the user or add-in requests them.
 
### Set calculation type

The [Application](/javascript/api/excel/excel.application) object provides a method to force an immediate recalculation. `Application.calculate(calculationType)` starts a manual recalculation based on the specified `calculationType`. The following values can be specified:

 - `full`: Recalculate all formulas in all open workbooks, regardless of whether they have changed since the last recalculation.
 - `fullRebuild`: Check dependent formulas, and then recalculate all formulas in all open workbooks, regardless of whether they have changed since the last recalculation.
 - `recalculate`: Recalculate formulas that have changed (or been programmatically marked for recalculation) since the last calculation, and formulas dependent on them, in all active workbooks.
 
> [!NOTE] 
> For more information about recalculation, see the [Change formula recalculation, iteration, or precision](https://support.office.com/article/change-formula-recalculation-iteration-or-precision-73fc7dac-91cf-4d36-86e8-67124f6bcce4) article.

### Temporarily suspend calculations

The Excel API also lets add-ins turn off calculations until `RequestContext.sync()` is called. This is done with `suspendApiCalculationUntilNextSync()`. Use this method when your add-in is editing large ranges without needing to access the data between edits.

```js
context.application.suspendApiCalculationUntilNextSync();
```

## See also

- [Fundamental programming concepts with the Excel JavaScript API](excel-add-ins-core-concepts.md)
- [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md)
- [Work with ranges using the Excel JavaScript API](excel-add-ins-ranges.md)