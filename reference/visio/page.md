# Page object (JavaScript API for Visio)

Applies to: _Visio Online_
>**Note:** The Visio JavaScript APIs are currently in preview and are subject to change. The Visio JavaScript APIs are not currently supported for use in production environments.

Represents the Page class.

## Properties

| Property	   | Type	|Description| Feedback|
|:---------------|:--------|:----------|:---|
|index|int|Index of the Page. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-page-index)|
|isBackground|bool|Whether the page is a background page or not. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-page-isBackground)|
|name|string|Page name. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-page-name)|

## Relationships
| Relationship | Type	|Description| Feedback|
|:---------------|:--------|:----------|:---|
|shapes|[ShapeCollection](shapecollection.md)|Shapes in the Page. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-page-shapes)|
|view|[PageView](pageview.md)|Returns the view of the page. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-page-view)|

## Methods

| Method		   | Return Type	|Description| Feedback|
|:---------------|:--------|:----------|:---|
|[activate()](#activate)|void|Set the page as Active Page of the document.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-page-activate)|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-page-load)|

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
