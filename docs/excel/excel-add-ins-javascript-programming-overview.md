# Excel JavaScript API programming overview

This article describes how to use the Excel JavaScript API to build add-ins for Excel 2016. It introduces key concepts that are fundamental to using the APIs, such as RequestContext, JavaScript proxy objects, sync(), Excel.run(), and load(). The code examples at the end of the article show you how to apply the concepts.

## RequestContext

The RequestContext object facilitates requests to the Excel application. Because the Office Add-in and the Excel application run in two different processes, request context is required to get access to Excel and related objects such as worksheets and tables, from the add-in. A request context is created as shown.

```js
var ctx = new Excel.RequestContext();
```

## Proxy objects

The Excel JavaScript objects declared and used in an add-in are proxy objects for the real objects in an Excel document. All actions taken on proxy objects are not realized in Excel, and the state of the Excel document is not realized in the proxy objects until the document state has been synchronized. The document state is synchronized when context.sync() is run (see below).

For example, the local JavaScript object `selectedRange` is declared to reference the selected range. This can be used to queue the setting of its properties and invoking methods. The actions on such objects are not realized until the sync() method is run.

```js
var selectedRange = ctx.workbook.getSelectedRange();
```

## sync()

The sync() method available on the request context synchronizes the state between JavaScript proxy objects and real objects in Excel by executing instructions queued on the context and retrieving properties of loaded Office objects for use in your code.  This method returns a promise, which is resolved when  synchronization is complete.

## Excel.run(function(context) { batch })

Excel.run() executes a batch script that performs actions on the Excel object model. The batch commands include definitions of local JavaScript proxy objects and sync() methods that synchronize the state between local and Excel objects and promise resolution. The advantage of batching requests in Excel.run() is that when the promise is resolved, any tracked range objects that were allocated during the execution will be automatically released.

The run method takes in RequestContext and returns a promise (typically, just the result of ctx.sync()). It is possible to run the batch operation outside of the Excel.run(). However, in such a scenario, any range object references needs to be manually tracked and managed.

## load()

The load() method is used to fill in the proxy objects created in the add-in JavaScript layer. When trying to retrieve an object, a worksheet for example, a local proxy object is created first in the JavaScript layer. Such an object can be used to queue the setting of its properties and invoking methods. However, for reading object properties or relations, the load() and sync() methods need to be invoked first. The load() method takes in the properties and relations that need to be loaded when the sync() method is called.

_Syntax:_

```js
object.load(string: properties);
//or
object.load(array: properties);
//or
object.load({loadOption});
```
Where,

* `properties` is the list of properties and/or relationship names to be loaded specified as comma-delimited strings or array of names. See .load() methods under each object for details.
* `loadOption` specifies an object that describes the selection, expansion, top, and skip options. See object load [options](../../reference/excel/loadoption.md) for details.

## Example: Write values from an array to a range object

The following example shows you how to write values from an array to a range object.

The Excel.run() contains a batch of instructions. As part of this batch, a proxy object is created that references a range (address A1:B2) on the active worksheet. The value of this proxy range object is set locally. In order to read the values back, the `text` property of the range is instructed to be loaded onto the proxy object. All these commands are queued and run when ctx.sync() is called. The sync() method returns a promise that can be used to chain it with other operations.

```js
// Run a batch operation against the Excel object model. Use the context argument to get access to the Excel document.
Excel.run(function (ctx) {

	// Create a proxy object for the sheet
	var sheet = ctx.workbook.worksheets.getActiveWorksheet();
	// Values to be updated
	var values = [
				 ["Type", "Estimate"],
				 ["Transportation", 1670]
				 ];
	// Create a proxy object for the range
	var range = sheet.getRange("A1:B2");

	// Assign array value to the proxy object's values property.
	range.values = values;

	// Synchronizes the state between JavaScript proxy objects and real objects in Excel by executing instructions queued on the context
	return ctx.sync().then(function() {
			console.log("Done");
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

## Example: Copy values

The following example shows how to copy the values from Range A1:A2 to B1:B2 of the active worksheet by using load() method on the range object.

```js
// Run a batch operation against the Excel object model. Use the context argument to get access to the Excel document.
Excel.run(function (ctx) {

	// Create a proxy object for the range
	var range = ctx.workbook.worksheets.getActiveWorksheet().getRange("A1:A2");

	// Synchronizes the state between JavaScript proxy objects and real objects in Excel by executing instructions queued on the context
	return ctx.sync().then(function() {
		// Assign the previously loaded values to the new range proxy object. The values will be updated once the following .then() function is invoked.
		ctx.workbook.worksheets.getActiveWorksheet().getRange("B1:B2").values = range.values;
	});
}).then(function() {
	  console.log("done");
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

## Properties and relationships selection

By default, object.load() selects all scalar and complex properties of the object that is being loaded. The relationships are not loaded by default (example, format is a relationship object of Range object). However, we recommend that you mark the properties and relations to be loaded explicitly to improve performance. To do this, specify (in the `load()` parameter) a subset of properties and relationships to include in the response. Load method allows two kinds of inputs:

* Property and relationship names as comma-separated string names _or_ as an array of strings containing property or relationship names.
* An object that describes the selection, expansion, top, and skip options. See object load [options](../../reference/excel/loadoption.md) for details.

```js
object.load  ('<var1>,<relation1/var2>');

// Pass the parameter as an array.
object.load (["var1", "relation1/var2"]);
```

### Example

The following load statement loads all the properties of the Range, and then expands on the format and format/fill.

```js
Excel.run(function (ctx) {
	var sheetName = "Sheet1";
	var rangeAddress = "A1:B2";
	var myRange = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);

	myRange.load(["address", "format/*", "format/fill", "entireRow" ]);
	return ctx.sync().then(function() {
		console.log (myRange.address); //ok
		console.log (myRange.format.wrapText); //ok
		console.log (myRange.format.fill.color); //ok
		//console.log (myRange.format.font.color); //not ok as it was not loaded

	});
}).then(function() {
	  console.log("done");
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

## Null-Input

### null input in 2-D Array

`null` input inside two-dimensional array (for values, number-format, formula) is ignored in the update API. No update will take place to the intended target when `null` input is sent in values or number-format or formula grid of values.

Example: In order to only update specific parts of the Range, such as some cell's Number Format, and to retain the existing number-format on other parts of the Range, set desired Number Format where needed and send `null` for the other cells.

In the following set request, only some parts of the Range Number Format are set while retaining the existing Number Format on the remaining part (by passing nulls).

```js
  range.values = [["Eurasia", "29.96", "0.25", "15-Feb" ]];
  range.numberFormat = [[null, null, null, "m/d/yyyy;@"]];
```
### null input for a property

`null` is not a valid single input for the entire property. For example, the following is not valid as the entire values cannot be set to null or ignored.

```js
 range.values= null;

```

The following is not valid either as null is not a valid color value.

```js
 range.format.fill.color =  null;
```

### Null-Response

Representation of formatting properties that consists of non-uniform values would result in the return of a null value in the response.

Example: A Range can consist of one of more cells. In cases where the individual cells contained in the Range specified don't have uniform formatting values, the range level representation will be undefined.

```js
  "size" : null,
  "color" : null,
```

### Blank input and output

Blank values in update requests are treated as instruction to clear or reset the respective property. Blank value is represented by two double quotation marks with no space in-between. `""`

Example:

* For `values`, the range value is cleared out. This is the same as clearing the contents in the application.

* For `numberFormat`, the number format is set to `General`.

* For `formula` and `formulaLocale`, the formula values are cleared.


For read operations, expect to receive blank values if the contents of the cells are blanks. If the cell contains no data or value, then the API returns a blank value. Blank value is represented by two double quotation marks with no space in-between. `""`.

```js
  range.values = [["", "some", "data", "in", "other", "cells", ""]];
```

```js
  range.formula = [["", "", "=Rand()"]];
```

## Unbounded range

### Read

Unbounded range address contains only column or row identifiers and unspecified row identifier or column identifiers (respectively), such as:

* `C:C`, `A:F`, `A:XFD` (contains unspecified rows)
* `2:2`, `1:4`, `1:1048546` (contains unspecified columns)

When the API makes a request to retrieve an unbounded Range (e.g., `getRange('C:C')`, the response returned contains `null` for cell level properties such as `values`, `text`, `numberFormat`, `formula`, etc.. Other Range properties such as `address`, `cellCount`, etc. will reflect the unbounded range.

### Write

Setting cell level properties (such as values, numberFormat, etc.) on unbounded Range is **not allowed** as the input request might be too large to handle.

Example: The following is not a valid update request because the requested range is unbounded.

```js
...
	var range = ctx.workbook.worksheets.getActiveWorksheet().getRange("A:B");
	range.values = 'Due Date';
...
```

When an update operation is attempted on such a Range, the API will return an error.


## Large range

Large range implies a Range whose size is too large for a single API call. Many factors such as number of cells, values, numberFormat, and formulas contained in the range can make the response so large that it becomes unsuitable for API interaction. The API makes a best attempt to return or write to the requested data. However, the large size involved might result in an API error condition because of the large resource utilization.

To avoid sthis, we recommend that you use read or write for large Range in multiple smaller range sizes.


## Single input copy

To support updating a range with the same values or number-format or applying same formula across a range, the following convention is used in the set API. In Excel, this behavior is similar to inputting values or formulas to a range in the CTRL+Enter mode.

The API will look for a *single cell value* and, if the target range dimension doesn't match the input range dimension, it will apply the update to the entire range in the CTRL+Enter model with the value or formula provided in the request.

### Examples

The following request updates the selected range with the text of "Due Date". Note that Range has 20 cells, whereas the provided input only has 1 cell value.

```js
Excel.run(function (ctx) {
	var sheetName = 'Sheet1';
	var rangeAddress = 'A1:A20';
	var worksheet = ctx.workbook.worksheets.getItem(sheetName);
	var range = worksheet.getRange(rangeAddress);
	range.values = 'Due Date';
	range.load('text');
	return ctx.sync().then(function() {
		console.log(range.text);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

The following request updates the selected range with the date of '3/11/2015'.

```js
Excel.run(function (ctx) {
	var sheetName = 'Sheet1';
	var rangeAddress = 'A1:A20';
	var worksheet = ctx.workbook.worksheets.getItem(sheetName);
	var range = worksheet.getRange(rangeAddress);
	range.numberFormat = 'm/d/yyyy';
	range.values = '3/11/2015';
	range.load('text');
	return ctx.sync().then(function() {
		console.log(range.text);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
The following request updates the selected range with a formula that will be applied across the range in the CTRL+Enter mode.

```js
Excel.run(function (ctx) {
	var sheetName = 'Sheet1';
	var rangeAddress = 'A1:A20';
	var worksheet = ctx.workbook.worksheets.getItem(sheetName);
	var range = worksheet.getRange(rangeAddress);
	range.numberFormat = 'm/d/yyyy';
	range.values = '3/11/2015';
	range.load('text');
	return ctx.sync().then(function() {
		console.log(range.text);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


## Error messages

Errors are returned using an error object that consists of a code and a message. The following table provides a list of possible error conditions that can occur.

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
|Conflict	|Request could not be processed because of a conflict.|
|ItemAlreadyExists|The resource being created already exists.|
|UnsupportedOperation|The operation being attempted is not supported.|
|RequestAborted|The request was aborted during run time.|
|ApiNotAvailable|The requested API is not available.|
|InsertDeleteConflict|The insert or delete operation attempted resulted in a conflict.|
|InvalidOperation|The operation attempted is invalid on the object.|

## Additional resources

* [Build your first Excel add-in](build-your-first-excel-add-in.md)
* [Code snippet explorer](https://github.com/OfficeDev/office-js-snippet-explorer)
* [Excel add-ins code samples](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [Excel add-ins JavaScript API reference](excel-add-ins-javascript-api-reference.md)
