# DocumentView object (JavaScript API for Visio)

Applies to: _Visio Online_
>**Note:** The Visio JavaScript APIs are currently in preview and are subject to change. The Visio JavaScript APIs are not currently supported for use in production environments.

Represents the DocumentView class.

## Properties

| Property	   | Type	|Description| Feedback|
|:---------------|:--------|:----------|:---|
|disableHyperlinks|bool|Disable Hyperlinks.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-documentView-disableHyperlinks)|
|disablePan|bool|Disable Pan.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-documentView-disablePan)|
|disableZoom|bool|Disable Zoom.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-documentView-disableZoom)|
|hideDiagramBoundry|bool|Disable Hyperlinks.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-documentView-hideDiagramBoundry)|

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Feedback|
|:---------------|:--------|:----------|:---|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-documentView-load)|

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
