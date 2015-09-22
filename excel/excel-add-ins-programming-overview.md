# Excel Add-ins JavaScript Programming Overview

_Applies to: Excel 2016, Office 2016_

This topic covers the fundamentals of using the JavaScript APIs to build the next generation of Excel add-ins. 


* [The basics](#the-basics)
* [Properties and Relations Selection](#properties-and-relations-selection)
* [Document Binding](#null-input)
* [Reference Binding](#null-input)
* [Null-Input](#null-input)
* [Null-Response](#null-response)
* [Blank Input and Output](#blank-input-and-output)
* [Unbounded-Range](#unbounded-range)
* [Large-Range](#large-range)
* [Single Input Copy](#single-input-copy)

## The basics

In this section, get a brief introduction to the key concepts that are fundamental to using the APIs, such as RequestContext, sync(), run(), and load().

### RequestContext
The RequestContext object facilitates requests to the Excel application. Since the Office add-in and the Excel application run in two different processes, request context is required to get access to Excel and related objects such as worksheets, tables, etc. from the add-in. 

### sync()

The sync() method available on the request context synchronizes the state between JavaScript proxy objects and real objects in Office by executing instructions queued on the context and retrieving properties of loaded Office objects for use in your code.  This method returns a promise, which is resolved when the synchronization is complete.

### Excel.run(function(context) {batch})

Executes a batch script that performs actions on the Excel object model. When the promise is resolved, any tracked objects that were automatically allocated during the execution will be released. 

Batch: A function that takes in RequestContext and returns a promise (typically, just the result of ctx.sync()). The RequestContext parameter facilitates requests to the Excel application. Since the Office add-in and the Excel application run in two different processes, the request context is required to get access to the Excel object model from the add-in.


#### Example

The following example shows how to write values from an array to a range. First, a RequestContext is created to get access to the workbook. Then, a worksheet is added. Range A1:B2 on the sheet is retrieved afterward. Finally, we assign the values stored in the array to this range. All these commands are queued and will run when ctx.sync() is called.  sync() returns a promise that can be used to chain it with other operations.

```js
Excel.run(function (ctx) { 
	var sheet = ctx.workbook.worksheets.add();
	var values = [
				 ["Type", "Estimate"],
				 ["Transportation", 1670]
				 ];
	var range = sheet.getRange("A1:B2");
	range.values = values;
	//Statements queued above will not be executed until the sync() is called. 
	return ctx.sync().then(function() {
			console.log("Done");
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
})
```

### load()
Load method is used to fill in the Excel proxy objects created in the add-in JavaScript layer. When trying to retrieve an object, for example a worksheet, a local proxy object is created first in the JavaScript layer. Such an object can be used to queue up setting of its properties and invoking methods. However, for reading object properties or relations, the load() method and sync() needs to be invoked first. Load method takes in the parameters and relations that need to be loaded when sync is called. 

#### Syntax

```js
object.load(properties);
//or
object.load({loadOption});
```
Where, 

* properties is the list of properties and/or relationship names to be loaded specified as comma delimited strings or array of names. See .load() methods under each object for details.
* loadOption specifies selection, expansion, top, and skip options. See [loadOption](resources/loadoption.md) object for details.

#### Example
The following example shows how to copy the values from Range A1:A2 to B1:B2 by using load() method on the range object.

```js
Excel.run(function (ctx) { 
	var range = ctx.workbook.worksheets.getActiveWorksheet().getRange("A1:A2");
	range.load ("address, values, range/format"); 
	// same as range.load (["address", "values", "range/format"]); 
	return ctx.sync().then(function() {
		var myvalues=range.values;
		ctx.workbook.worksheets. getActiveWorksheet().getRange("B1:B2").values= myvalues;
	});
}).then(function() {
	  console.log("done");
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
})
```

### Summary
1.	Getting a RequestContext is the first step to interact with Excel.
2.	All JavaScript objects are local proxy objects.  Any method invocation or setting of properties, queues up the commands in JavaScript, but does not submit them until sync() is called. 
3.	Load is a special type of command for retrieval of properties. Properties can only be accessed after invoking sync(). 
4.	For performance reasons, avoid loading objects without specifying individual properties that will be used.

[top](#programming-notes)

## Properties and Relations Selection 

In this and the upcoming sections, learn about the other core concepts that are important to using these APIs.

* By default load() selects all scalar/complex properties of the object which is being loaded. The relations are not loaded by default.  Exceptions:  any binary, XML, etc. properties are not returned. 
* The select option specifies a subset of properties and/or relations to include in the response.
* The properties to be selected are provided during the load statement.
* Select will essentially get the users into optimized mode of handpicking what they want. 
* Property names are listed as a parameter to the select property. Support two kinds of inputs
	* Property names are separated by comma. 
	* Provide an array of property name strings

```js	
object.load  (<var1>,<relation1/var2>);

// Pass the parameter as an array.
object.load (["var1", "relation1/var2"]);
```

#### Examples

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:B2";
	var myRange = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	//load statement below loads the address, values, and numberFormat properties of the Range and then expands on the format, format/background, entireRow relations
	myRange.load (["address", "values", "numberFormat", "format", "format/background", "entireRow"]);
	return ctx.sync().then(function() {
		console.log (myRange.address); //ok
		console.log (myRange.cellCount); //not-ok
		console.log (myRange.format.wrapText); //ok
		console.log (myRange.format.background.color); //ok
		console.log (myRange.format.font.color); //not-ok
		console.log (myRange.entireRow.address); //ok
		console.log (myRange.entireColumn.address); //not-ok
	});
}).then(function() {
	  console.log("done");
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
})	
```
Load statement below loads all the properties of the Range and then expands on the format, format/background, entireRow relations.  
 
```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:B2"; 
	myRange.load(["address", "format", "format/background", "entireRow" ]);
	return ctx.sync().then(function() {
		console.log (myRange.address); //ok
		console.log (myRange.format.wrapText); //ok
		console.log (myRange.format.background.color); //ok
		console.log (myRange.format.font.color); //not-ok
		console.log (myRange.entireRow.address); //ok
		console.log (myRange.entireColumn.address); //not-ok
	});
}).then(function() {
	  console.log("done");
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
})
```



## Null-Input

### null input in 2-D Array

`null` input inside two dimensional array (for values, number-format, formula) is ignored in the update API. No update will take place to the intended target when `null` input is sent in values or number-format or formula grid of values.

Example: In order to only update specific parts of the Range such as some cell's Number Format and retain the existing number-format on other parts of the Range, set desired Number Format where needed and send `null` for the other cells. 

In the set request below, only some parts of the Range Number Format is set while retaining the existing Number Format on the remaining part (by passing nulls).

```js
  range.values = [["Eurasia", "29.96", "0.25", "15-Feb" ]];
  range.numberFormat = [[null, null, null, "m/d/yyyy;@"]];
```
#### null input for a property

`null` is not a valid single input for the entire property. For example, the following is not valid as the entire values cannot be set to null or ignored. 

```
 range.values= null;

```

Following is not valid either as null is not a valid color value. 
```
 range.format.background.color =  null;
```

## Null-Response

Representation of formatting properties that consists of non-uniform values would result in `null` value to be returned in the response. 

Example: A Range can consist of one of more cells. In cases where the individual cells contained in the Range specified don't have uniform formatting values, the range level representation will be undefined. 

```
  "size" : null,
  "color" : null,
```

## Blank Input and Output

Blank values in update requests are treated as instruction to clear or reset the respective property. Blank value is represented by two double-quotes with no space in between. `""`

Example: 
* For `values`, the range value is cleared out. This is same as clearing the contents in the application.
* For `numberFormat`, the number format is set to `General`.
* For `formula` and `formulaLocale`, the formula values are cleared. 

For read operations, expect to receive blank values if the contents of the cells are blanks. If the cell contains no data or value, then the API returns a blank value. Blank value is represented by two double quotation marks with no space in between. `""`.

```
  range.values = [["", "some", "data", "in", "other", "cells", ""]];
```

```
  range.formula = [["", "", "=Rand()"]];
```

## Unbounded-Range

### Read

Unbounded range address contains only column or row identifiers and unspecified row identifier or column identifiers (respectively), such as:

* `C:C`, `A:F`, `A:XFD` (contains unspecified rows)
* `2:2`, `1:4`, `1:1048546` (contains unspecified columns)

When the API makes a request to retrieve an unbounded Range (e.g., `getRange('C:C')`, the response returned contains `null` for cell level properties such as `values`, `text`, `numberFormat`, `formula`, etc.. Other Range properties such as `address`, `cellCount`, etc. will reflect the unbounded range.

### Write

Setting cell level properties (such as values, numberFormat, etc.) on unbounded Range is **not allowed** as the input request might be too large to handle. 

Example: The following is not a valid update request because the requested range is unbounded. 

```js
Excel.run(function (ctx) { 
	var sheetName = 'Sheet1';
	var rangeAddress = 'A:B';
	var ctx = new Excel.ExcelClientContext();
	var worksheet = ctx.workbook.worksheets.getItem(sheetName);
	var range = worksheet.getRange(rangeAddress);
	range.values = 'Due Date';
	range.load('text');
	return ctx.sync().then(function() {
		Console.log(range.text);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
})
```

When such a Range is update operation is attempted, the API returns the an error.


## Large-Range

Large Range implies a Range whose size is too large for a single API call. Many factors such as number of cells,values, numberFormat, formulas, etc. contained in the range can make the response so large, it becomes unsuitable for API interaction. The API makes best attempt to return or write-to the requested data. However, large size involved might result in API error condition because of the large resource utilization. 

To avoid such a condition, using read or write for large Range in multiple smaller range sizes is recommended .


## Single Input Copy

To support updating a range with same values or number-format or applying same formula across a range, the following convention is used in the set API. In Excel, this behavior is similar to inputting values or formulas to a range in the CTRL+Enter mode. 

API will look for *single cell value* and if the target range dimension doesn't match the input range dimension it will apply the update to the entire range in the CTRL+Enter model with the value or formula provided in the request.

### Examples

Following request updates selected range with the a text of "Due Date". Note that Range has 20 cells whereas the provided input only has 1 cell value.

```js
Excel.run(function (ctx) { 
	var sheetName = 'Sheet1';
	var rangeAddress = 'A1:A20';
	var ctx = new Excel.ExcelClientContext();
	var worksheet = ctx.workbook.worksheets.getItem(sheetName);
	var range = worksheet.getRange(rangeAddress);
	range.values = 'Due Date';
	range.load('text');
	return ctx.sync().then(function() {
		Console.log(range.text);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
})
```

Following request updates selected range with date of '3/11/2015'.


```js
Excel.run(function (ctx) { 
	var sheetName = 'Sheet1';
	var rangeAddress = 'A1:A20';
	var ctx = new Excel.ExcelClientContext();
	var worksheet = ctx.workbook.worksheets.getItem(sheetName);
	var range = worksheet.getRange(rangeAddress);
	range.numberFormat = 'm/d/yyyy';
	range.values = '3/11/2015';
	range.load('text');
	return ctx.sync().then(function() {
		Console.log(range.text);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
})
```
Following request updates selected range with a formula of that will be applied across in the CTRL+Enter mode.  

```js
Excel.run(function (ctx) { 
	var sheetName = 'Sheet1';
	var rangeAddress = 'A1:A20';
	var ctx = new Excel.ExcelClientContext();
	var worksheet = ctx.workbook.worksheets.getItem(sheetName);
	var range = worksheet.getRange(rangeAddress);
	range.formula = '=DAYS(B15,42060)';
	range.load'(text');
	return ctx.sync().then(function() {
		Console.log(range.text);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
})
```


## Error Messages

Errors are returned using an error object that consists of a code and a message. The following table provides a list of possible error conditions that can occur in the add-ins. 

|error.code | error.message |
|:----------|:--------------|
|InvalidArgument |The argument is invalid or missing or has an incorrect format.|
|InvalidRequest  |Cannot process the request.|
|InvalidReference|This reference is not valid for the current operation.|
|InvalidBinding  |This object binding is no longer valid due to previous updates.|
|InvalidSelection|The current selection is invalid for this operation.|
|Unauthenticated |Required authentication information is either missing or invalid.|
|AccessDenied	|You cannot perform the requested operation.|
|ItemNotFound	|The requested resource doesn't exist.|
|ActivityLimitReached|Activity limit has been reached.|
|GeneralException|There was an internal error while processing the request.|
|NotImplemented  |The requested feature isn't implemented.|
|ServiceNotAvailable|The service is unavailable.|
|Conflict	|Request could not be processed because of conflict.|
|ItemAlreadyExists|The resource being created already exists.|
|UnsupportedOperation|The operation being attempted is not supported.|
|RequestAborted|The request was aborted during run time.|
|ApiNotAvailable|The requested API is not available.|
|InsertDeleteConflict|The insert or delete operation attempted resulted in conflict.|
|InvalidOperation|The operation attempted is invalid on the object.|


## Learn more

The following are just a few of the available resources. 

1.  [Build your first Excel Add-in](build-your-first-excel-add-in.md)
2.  [Snippet Explorer for Excel](http://officesnippetexplorer.azurewebsites.net/#/snippets/excel)
3.  [Excel Add-ins code samples](excel-add-ins-code-samples.md) 
4.  [Excel Add-ins JavaScript API Reference](excel-add-ins-javascript-reference.md)