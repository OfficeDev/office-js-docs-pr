# PageView object (JavaScript API for Visio)

Applies to: _Visio Online_

Represents the PageView class.

## Properties

| Property | Type |Description|
|:---------------|:--------|:----------|
|zoom|int|GetSet Page's Zoom level.|

## Relationships
None

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[centerViewportOnShape(ShapeId: number)](#centerviewportonshapeshapeid-number)|void|Pans the Visio drawing to place the specified shape in the center of the view.|
|[fitToWindow()](#fittowindow)|void|Fit Page to current window.|
|[getPosition()](#getposition)|[Position](position.md)|Returns the position object that specifies the position of the page in the view.|
|[getSelection()](#getselection)|[Selection](selection.md)|Represents the Selection in the page.|
|[isShapeInViewport(Shape: Shape)](#isshapeinviewportshape-shape)|bool|To check if the shape is in view of the page or not.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|
|[setPosition(Position: Position)](#setpositionposition-position)|void|Sets the position of the page in the view.|

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

### getPosition()
Returns the position object that specifies the position of the page in the view.

#### Syntax
```js
pageViewObject.getPosition();
```

#### Parameters
None

#### Returns
[Position](position.md)

### getSelection()
Represents the Selection in the page.

#### Syntax
```js
pageViewObject.getSelection();
```

#### Parameters
None

#### Returns
[Selection](selection.md)

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

### setPosition(Position: Position)
Sets the position of the page in the view.

#### Syntax
```js
pageViewObject.setPosition(Position);
```

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

