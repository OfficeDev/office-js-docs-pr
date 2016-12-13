# ShapeCollection object (JavaScript API for Visio)

_Visio Online_

Represents the Shape Collection.

## Properties

| Property	   | Type	|Description| Req. Set| Feedback|
|:---------------|:--------|:----------|:----|:---|
|items|[Shape[]](shape.md)|A collection of shape objects. Read-only.|1.1|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeCollection-items)|

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set| Feedback|
|:---------------|:--------|:----------|:----|:---|
|[getCount()](#getcount)|int|Gets the number of Shapes in the collection.|1.1|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeCollection-getCount)|
|[getItem(key: number or string)](#getitemkey-number-or-string)|[Shape](shape.md)|Gets a Shape using its key (name or Index).|1.1|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeCollection-getItem)|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.1|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeCollection-load)|

## Method Details


### getCount()
Gets the number of Shapes in the collection.

#### Syntax
```js
shapeCollectionObject.getCount();
```

#### Parameters
None

#### Returns
int

#### Examples
```js
Visio.run(function (ctx) { 
	var activePage = ctx.document.getActivePage();
	var numShapesActivePage = activePage.shapes.getCount();
	return ctx.sync().then(function () {
		console.log("Shapes Count: " + numShapesActivePage.value);
	});

}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getItem(key: number or string)
Gets a Shape using its key (name or Index).

#### Syntax
```js
shapeCollectionObject.getItem(key);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|key|number or string|Key is the Name or Index of the shape to be retrieved.|

#### Returns
[Shape](shape.md)

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
