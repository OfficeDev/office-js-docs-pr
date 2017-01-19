# ShapeView object (JavaScript API for Visio)

Applies to: _Visio Online_
>**Note:** The Visio JavaScript APIs are not currently available for use in preview or production environments.

Represents the ShapeView class.

## Properties

None

## Relationships
None

## Methods

| Method		   | Return Type	|Description| Feedback|
|:---------------|:--------|:----------|:---|
|[addOverlay(OverlayType: OverlayType, Content: string, HorizontalAlignment: HorizontalAlignment, VerticalAlignment: VerticalAlignment, Width: number, Height: number)](#addoverlayoverlaytype-overlaytype-content-string-horizontalalignment-horizontalalignment-verticalalignment-verticalalignment-width-number-height-number)|int|Adds an overlay on top of the shape.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeView-addOverlay)|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeView-load)|
|[removeOverlay(OverlayId: number)](#removeoverlayoverlayid-number)|void|Removes particular overlay or all overlays on the Shape.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeView-removeOverlay)|

## Method Details


### addOverlay(OverlayType: OverlayType, Content: string, HorizontalAlignment: HorizontalAlignment, VerticalAlignment: VerticalAlignment, Width: number, Height: number)
Adds an overlay on top of the shape.

#### Syntax
```js
shapeViewObject.addOverlay(OverlayType, Content, HorizontalAlignment, VerticalAlignment, Width, Height);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|OverlayType|OverlayType|An Overlay Type -Text, Image.|
|Content|string|Content of Overlay.|
|HorizontalAlignment|HorizontalAlignment|Horizontal Alignment of Overlay - Left, Center, Right|
|VerticalAlignment|VerticalAlignment|Vertical Alignment of Overlay - Top, Middle, Bottom|
|Width|number|Overlay Width.|
|Height|number|Overlay Height.|

#### Returns
int

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

### removeOverlay(OverlayId: number)
Removes particular overlay or all overlays on the Shape.

#### Syntax
```js
shapeViewObject.removeOverlay(OverlayId);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|OverlayId|number|An Overlay Id. Removes the specific overlay id from the shape.|

#### Returns
void

### Property access examples
```js
Visio.run(function (ctx) { 
	var activePage = ctx.document.getActivePage();
	var shape = activePage.shapes.getItem(0);
	var overlayId=shape.view.addOverlay(1, "Visio Online", 2, 2, 50, 50);
	return ctx.sync();
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
	shape.view.removeOverlay(1);
	return ctx.sync();
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
