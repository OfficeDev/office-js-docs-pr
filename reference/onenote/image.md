# Image Object (JavaScript API for OneNote)

_Applies to: OneNote Online_
_Note: This API is in preview_


Represents an image.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|description|string|Gets the description of the image. Read-only.|
|height|double|Gets the height of the image layout. Read-only.|
|hyperlink|string|Gets or sets the hyperlink of the image.|
|id|string|Gets the ID of the image. Read-only.|
|width|double|Gets or sets the width of the image layout.|


## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|pageContent|[PageContent](pagecontent.md)|Gets the page content object that contains the image, if the image is at the page level. Read-only.|
|paragraph|[Paragraph](paragraph.md)|Gets the paragraph that contains the image. Returns null if the image is in an outline. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[changeDescription(description: string)](#changedescriptiondescription-string)|void|Changes the description of the image.|
|[getBase64Image()](#getbase64image)|string|Gets the base64-encoded binary representation of the image.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### changeDescription(description: string)
Changes the description of the image.

#### Syntax
```js
imageObject.changeDescription(description);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|description|string|The description for the image.|

#### Returns
void

### getBase64Image()
Gets the base64-encoded binary representation of the image.

#### Syntax
```js
imageObject.getBase64Image();
```

#### Parameters
None

#### Returns
string

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
