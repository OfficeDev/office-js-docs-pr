---
title: Work with workbooks using the Excel JavaScript API
description: ''
ms.date: 11/15/2018
---


# Work with Workbooks using the Excel JavaScript API

This article provides code samples that show how to perform common tasks with workbooks using the Excel JavaScript API. For the complete list of properties and methods that the **Workbook** object supports, see [Workbook Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.workbook).

The workbook object is the entry point for your add-in to interact with Excel. It contains a [WorksheetCollection](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection) to manipulate individual worksheets. Most work done to [ranges](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-ranges) happens directly through a worksheet. The article [Work with Worksheets using the Excel JavaScript API](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-worksheets) describes how to access and edit worksheets.

## Get active cell and selected range

Workbooks maintain collections of worksheets, tables, PivotTables, and other data types through which Excel data is accessed and changed. They also have application-level information (more on that later). Their are a pair of methods to directly get ranges the Excel user has just activated. The first is `getActiveCell()`. This method gets the active cell from the workbook as a [Range object](https://docs.microsoft.com/javascript/api/excel/excel.range). The following example shows a call to `getActiveCell()`, followed by the cell's address being printed to the console.


```js
Excel.run(function (context) {
    var activeCell = context.workbook.getActiveCell();
	activeCell.load("address");

    return context.sync().then(function () {
		console.log("The active cell is " + activeCell.address);
	});
}).catch(errorHandlerFunction);
```

The second range access method is `getSelectedRange()`. This method returns the range currently selected by either a user or an add-in. The following example shows a call to `getSelectedRange()` that then sets the range's fill color to yellow.

```js
Excel.run(function(context) {
	var range = context.workbook.getSelectedRange();
	range.format.fill.color = "yellow";
	return context.sync();
}).catch(errorHandlerFunction);
```

## Create Workbook

Your add-in can create a new workbook, separate from the Excel instance in which the add-in is currently running. The Excel object has the `createWorkbook` method for this purpose. When this method is called, the new workbook is immediately opened and displayed. This happens in a new Excel instance. Your add-in remains open and running with the previous workbook.

```js
Excel.createWorkbook();
```

`createWorkbook` takes in an optional string. This string is an .xlsx file in base64 encoding. Assuming the string argument is a valid .xlsx file, the resulting workbook will be a copy of that file. You can get your add-in’s current workbook as a base64-encoded string by using [file slicing](https://docs.microsoft.com/javascript/api/office/office.document#getfileasync-filetype--options--callback-). The [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) class can be used to convert a file into the required base64-encoded string, as demonstrated in the following example. 

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

## Protection

Your add-in can control a user's ability to edit the worksheet structure. The workbook's `protection` property is a [WorkbookProtection](https://docs.microsoft.com/javascript/api/excel/excel.workbookprotection) object with a `protect()` method. The following example shows a basic scenario toggling the protection of the workbook's structure. 

```js
Excel.run(function (context) {
	var workbook = context.workbook;
	workbook.load("protection/protected");

	return context.sync().then(function() {
		if (!workbook.protection.protected) {
			workbook.protection.protect();
		}
	})
}).catch(errorHandlerFunction);
```

The `protect` method also has an optional string parameter, `password`. This string represents the password needed for a user to bypass protection and change the workbook's structure.

## Document properties

A workbook has access to the Office file metadata, which is known as the [document properties](https://support.office.com/article/View-or-change-the-properties-for-an-Office-file-21D604C2-481E-4379-8E54-1DD4622C6B75). The workbook's `properties` property is a [DocumentProperties](https://docs.microsoft.com/javascript/api/excel/excel.documentproperties) object containing these metadata values. The following example shows how to set the author property.

```js
Excel.run(function (context) {
	var docProperties = context.workbook.properties;
	docProperties.author = "Alex";
	return context.sync();
}).catch(errorHandlerFunction);
```

You can also define custom properties. The `DocumentProperties` object has `custom`: a map defining key-value pairs for user-defined properties. 

## Document settings



## Calculations

## See also

- [Fundamental programming concepts with the Excel JavaScript API](excel-add-ins-core-concepts.md)
- [Work with Worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md)
- [Work with ranges using the Excel JavaScript API](excel-add-ins-ranges.md)