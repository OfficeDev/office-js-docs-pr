# Hyperlink object (JavaScript API for Visio)

Applies to: _Visio Online_

Represents the Hyperlink.

## Properties

| Property	   | Type	|Description|
|:---------------|:--------|:----------|
|address|string|Gets the address of the Hyperlink object. Read-only.|
|description|string|Gets the description of a hyperlink. Read-only.|
|subAddress|string|Gets the sub-address of the Hyperlink object. Read-only.|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### load(param: object)
Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.

#### Syntax
```js
object.load(param);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|param|object|Optional. Accepts parameter and relationship names as delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns
void
### Property access examples
```js
Visio.run(function (ctx) { 
	var activePage = ctx.document.getActivePage();
	var shape = activePage.shapes.getItem(0);
	var hyperlink = shape.hyperlinks.getItem(0);
	hyperlink.load();
	return ctx.sync().then(function() {
		console.log(hyperlink.description);
		console.log(hyperlink.address);
		console.log(hyperlink.subAddress);
 	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
