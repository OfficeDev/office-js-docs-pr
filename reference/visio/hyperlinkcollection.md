# HyperlinkCollection object (JavaScript API for Visio)

Applies to: _Visio Online_
>**Note:** The Visio JavaScript APIs are currently in preview and are subject to change. The Visio JavaScript APIs are not currently supported for use in production environments.

Represents the Hyperlink Collection.

## Properties

| Property	   | Type	|Description| Req. Set| Feedback|
|:---------------|:--------|:----------|:----|:---|
|items|[Hyperlink[]](hyperlink.md)|A collection of hyperlink objects. Read-only.|1.1|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-hyperlinkCollection-items)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set| Feedback|
|:---------------|:--------|:----------|:----|:---|
|[getCount()](#getcount)|int|Gets the number of hyperlinks.|1.1|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-hyperlinkCollection-getCount)|
|[getItem(Key: number or string)](#getitemkey-number-or-string)|[Hyperlink](hyperlink.md)|Gets a Hyperlink using its key (name or Id).|1.1|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-hyperlinkCollection-getItem)|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.1|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-hyperlinkCollection-load)|

## Method Details


### getCount()
Gets the number of hyperlinks.

#### Syntax
```js
hyperlinkCollectionObject.getCount();
```

#### Parameters
None

#### Returns
int

### getItem(Key: number or string)
Gets a Hyperlink using its key (name or Id).

#### Syntax
```js
hyperlinkCollectionObject.getItem(Key);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|Key|number or string|Key is the name or index of the Hyperlink to be retrieved.|

#### Returns
[Hyperlink](hyperlink.md)

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
	var shapeName = "Manager Belt";
	var shape = activePage.shapes.getItem(shapeName);
	var hyperlinks = shape.hyperlinks;
	shapeHyperlinks.load();
		ctx.sync().then(function () {
			for(var i=0; i<shapeHyperlinks.items.length;i++)
				{
				  var hyperlink = shapeHyperlinks.items[i];
				  console.log("Description:"+hyperlink.description +"Address:"+hyperlink.address +"SubAddress:  "+ hyperlink.subAddress);
				}

			});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
