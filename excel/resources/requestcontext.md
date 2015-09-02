# RequestContext

The RequestContext object facilitates requests to the Excel application. Since the Office add-in and the Excel application run in two different processes, request context is required to get access to Excel and related objects such as worksheets, tables, etc. from the add-in. 

## Properties
None

## Methods

| Method         | Return Type    |Description|Notes |
|:---------------|:--------|:----------|:-----|
|[load(object: object, option: object)](#loadobject-object-option-object)  |void     |Fills the proxy object created in JavaScript layer with property and options specified in the parameter.||
|[executeAsync()](#executeasync)  |Promise Object |Submits the request queue to Excel and returns a promise object, which can be used for chaining further actions.||

## API Specification

### load(object: object, option: object)
Fills the proxy object created in JavaScript layer with property and options specified in the parameter.

#### Syntax
```js
requestContextObject.load(object, loadOption);
```

#### Parameters
| Parameter       | Type    |Description|
|:----------------|:--------|:----------|
|object|object|Optional. Specify the name of the object to be loaded.|
|option|[loadOption](loadoption.md)|Optional. Specify the load options such as select, expand, skip and top. Se Load Option object for details.|

#### Returns
void

##### Examples

The following example shows how to read and copy the values from Range A1:A2 to B1:B2.

```js
var ctx = new Excel.RequestContext();
var range = ctx.workbook.worksheets.getActiveWorksheet().getRange("A1:A2");
ctx.load(range, {"select": "address, values", "expand" : "range/format"});

ctx.executeAsync()
	.then(function () {
	var myvalues=range.values;
	ctx.workbook.worksheets. getActiveWorksheet().getRange("B1:B2").values= myvalues;
	ctx.executeAsync()
  		.then(function () {
			console.log(range.address);
			console.log(range.values);
			console.log(range.format.wrapText);
		})
		.catch(function(error) {
			console. error(JSON.stringify(error));
		})
});
```

### executeAsync() 
Promise Object |Submits the request queue to Excel and returns a promise object, which can be used for chaining further actions.

#### Syntax
```js
requestContextObject.executeAsync();
```

#### Parameters
None

#### Returns
Promise object.

##### Examples


```js
	var ctx = new Excel.RequestContext();
	var sheet = ctx.workbook.worksheets.add();

	ctx.executeAsync()
		.then(function () {   			
			console.log("Done");
		 })
		.catch(function(error) {
			console. error(JSON.stringify(error));
		});
```
