# Page object (JavaScript API for Visio)

Applies to: _Visio Online_

Represents the Page class.

## Properties

| Property	   | Type	|Description|
|:---------------|:--------|:----------|
|height|int|Returns the height of the page. Read-only.|
|index|int|Index of the Page. Read-only.|
|isBackground|bool|Whether the page is a background page or not. Read-only.|
|name|string|Page name. Read-only.|
|width|int|Returns the width of the page. Read-only.|

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|comments|[CommentCollection](commentcollection.md)|Returns the Comments Collection. Read-only.|
|allShapes|[ShapeCollection](shapecollection.md)|All shapes in the Page, including subshapes. Read-only.|
|shapes|[ShapeCollection](shapecollection.md)|All top-level shapes in the Page. Read-only.|
|view|[PageView](pageview.md)|Returns the view of the page. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[activate()](#activate)|void|Set the page as Active Page of the document.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### activate()
Set the page as Active Page of the document.

#### Syntax
```js
pageObject.activate();
```

#### Parameters
None

#### Returns
void

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
