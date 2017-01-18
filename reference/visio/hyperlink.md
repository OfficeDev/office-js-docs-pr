# Hyperlink object (JavaScript API for Visio)

Applies to: _Visio Online_
>**Note:** The Visio JavaScript APIs are not currently available for use in preview or production environments.

Represents the Hyperlink.

## Properties

| Property	   | Type	|Description| Feedback|
|:---------------|:--------|:----------|:---|
|address|string|Gets the address of the Hyperlink object. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-hyperlink-address)|
|description|string|Gets the description of a hyperlink. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-hyperlink-description)|
|subAddress|string|Gets the sub-address of the Hyperlink object. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-hyperlink-subAddress)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Feedback|
|:---------------|:--------|:----------|:---|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-hyperlink-load)|

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
