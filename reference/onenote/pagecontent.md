# PageContent Object (JavaScript API for OneNote)

_Applies to: OneNote Online_
_Note: This API is in preview_

Represents a region on a page that contains top-level content types such as Outline or Image. A PageContent object can be assigned an XY position.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|id|string|Gets the ID of the PageContent object. Read-only.|
|left|double|Gets or sets the left (X-axis) position of the PageContent object.|
|top|double|Gets or sets the top (Y-axis) position of the PageContent object.|
|type|string|Gets the type of the PageContent object. Read-only. Possible values are: Outline, Image, Other.|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|image|[Image](image.md)|Gets the Image in the PageContent object. Returns null if PageContentType is not Image. Read-only.|
|outline|[Outline](outline.md)|Gets the Outline in the PageContent object. Returns null if PageContentType is not Outline. Read-only.|
|parentPage|[Page](page.md)|Gets the page that contains the PageContent object. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[delete()](#delete)|void|Deletes the PageContent object.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|
|[select()](#select)|void|Selects the PageContent object.|

## Method Details


### delete()
Deletes the PageContent object.

#### Syntax
```js
pageContentObject.delete();
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
|:---------------|:--------|:----------|
|param|object|Optional. Accepts parameter and relationship names as delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns
void

### select()
Selects the PageContent object.

#### Syntax
```js
pageContentObject.select();
```

#### Parameters
None

#### Returns
void
