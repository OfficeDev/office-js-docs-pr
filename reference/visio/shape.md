# Shape object (JavaScript API for Visio)

Applies to: _Visio Online_

Represents the Shape class.

## Properties

| Property	   | Type	|Description|
|:---------------|:--------|:----------|
|id|int|Shape's Identifier. Read-only.|
|name|string|Shape's name. Read-only.|
|select|bool|Returns true, if shape is selected. User can set true to select the shape explicitly.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shape-select)|
|text|string|Shape's Text. Read-only.|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|comments|[CommentCollection](commentcollection.md)|Returns the Comments Collection. Read-only.|
|hyperlinks|[HyperlinkCollection](hyperlinkcollection.md)|Returns the Hyperlinks collection for a Shape object. Read-only.|
|shapeDataItems|[ShapeDataItemCollection](shapedataitemcollection.md)|Returns the Shape's Data Section. Read-only.|
|subShapes|[ShapeCollection](shapecollection.md)|Gets SubShape Collection. Read-only.|
|view|[ShapeView](shapeview.md)|Returns the view of the shape. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getBounds()](#getbounds)|[BoundingBox](boundingbox.md)|Returns the BoundingBox object that specifies bounding box of the shape.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### getBounds()
Returns the BoundingBox object that specifies bounding box of the shape.

#### Syntax
```js
shapeObject.getBounds();
```

#### Parameters
None

#### Returns
[BoundingBox](boundingbox.md)

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
	var shapeName = "Sample Name";
	var shape = activePage.shapes.getItem(shapeName);
	shape.load();
	return ctx.sync().then(function () {
		console.log(shape.name );
		console.log(shape.id );
		console.log(shape.Text );
		console.log(shape.Select );
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### Property access examples
```js
Visio.run(function (ctx) { 
	var activePage = ctx.document.getActivePage();
	var shape = activePage.shapes.getItem(0);
	shape.view.highlight = { color: "#E7E7E7", width: 100 };
	return ctx.sync();
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
