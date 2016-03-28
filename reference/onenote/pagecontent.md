# PageContent Object (JavaScript API for OneNote)

_Applies to: OneNote Online_
_Note: This API is in public preview_

Represents a placeholder object for the top-level content objects of a page. Top-level content objects can be positioned on the page.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|id|string|Gets the ID of the page content object. Read-only.|
|left|double|Gets or sets the left (x) position of the page content object.|
|top|double|Gets or sets the top (y) position of the page content object.|
|type|string|Gets the type of page content. Read-only. Possible values are: Outline, Image, Ink, InsertedFile, MediaFile, Other.|

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|image|[Image](image.md)|Gets the image in the page content object. Returns null if PageContentType is not Image. Read-only.|
|outline|[Outline](outline.md)|Gets the outline in the page content object. Returns null if PageContentType is not Outline. Read-only.|
|page|[Page](page.md)|Gets the page that contains the page content object. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[delete()](#delete)|void|Deletes the page content object.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### delete()
Deletes the page content object.

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
