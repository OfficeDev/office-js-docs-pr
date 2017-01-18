# PageView object (JavaScript API for Visio)

Applies to: _Visio Online_
>**Note:** The Visio JavaScript APIs are not currently available for use in preview or production environments.

Represents the PageView class.

## Properties

| Property | Type |Description| Feedback|
|:---------------|:--------|:----------|:---|
|zoom|int|GetSet Page's Zoom level.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-pageView-zoom)|

## Relationships

None

## Methods

| Method		   | Return Type	|Description| Feedback|
|:---------------|:--------|:----------|:---|
|[centerViewportOnShape(ShapeId: number)](#centerviewportonshapeshapeid-number)|void|Pans the Visio drawing to place the specified shape in the center of the view.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-pageView-centerViewportOnShape)|
|[fitToWindow()](#fittowindow)|void|Fit Page to current window.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-pageView-fitToWindow)|
|[isShapeInViewport(Shape: Shape)](#isshapeinviewportshape-shape)|bool|To check if the shape is in view of the page or not.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-pageView-isShapeInViewport)|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-pageView-load)|

## Method Details


### centerViewportOnShape(ShapeId: number)
Pans the Visio drawing to place the specified shape in the center of the view.

#### Syntax
```js
pageViewObject.centerViewportOnShape(ShapeId);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|ShapeId|number|ShapeId to be seen in the center.|

#### Returns
void

#### Examples
```js
Visio.run(function (ctx) { 
	var activePage = ctx.document.getActivePage();
	var shape = activePage.shapes.getItem(0);
	activePage.view.centerViewportOnShape(shape.Id);
	return ctx.sync();
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### fitToWindow()
Fit Page to current window.

#### Syntax
```js
pageViewObject.fitToWindow();
```

#### Parameters
None

#### Returns
void

### isShapeInViewport(Shape: Shape)
To check if the shape is in view of the page or not.

#### Syntax
```js
pageViewObject.isShapeInViewport(Shape);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|Shape|Shape|Shape to be checked.|

#### Returns
bool

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

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|Position|Position|Position object that specifies the new position of the page in the view.|

#### Returns
void
### Property access examples
```js
Visio.run(function (ctx) { 
	var activePage = ctx.document.getActivePage();
	activePage.view.zoom = 300;
	return ctx.sync();
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

