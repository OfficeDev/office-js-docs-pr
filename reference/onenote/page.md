# Page Object (JavaScript API for OneNote)

_Applies to: OneNote Online_
_Note: This API is in public preview_

Represents a OneNote page.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|id|string|Gets the ID of the page. Read-only.|
|pageLevel|int|Gets or sets the indentation level of page.|
|title|string|Gets the title of the page.|


## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|section|[Section](section.md)|Gets the section that contains the page. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[addImageFromBase64(left: double, top: double, base64EncodedImage: String)](#addimagefrombase64left-double-top-double-base64encodedimage-string)|[Image](image.md)|Adds an image to the page at the specified position. The image is added to the page's collection of page content objects.|
|[addOutline(left: double, top: double, html: String)](#addoutlineleft-double-top-double-html-string)|[Outline](outline.md)|Adds an outline to the page at the specified position. The outline is added to the page's collection of page content objects.|
|[getContents()](#getcontents)|[PageContentCollection](pagecontentcollection.md)|Gets the collection of page content objects for the page.|
|[insertPageAsSibling(location: string, title: string)](#insertpageassiblinglocation-string-title-string)|[Page](page.md)|Inserts a new page as a sibling of the page.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### addImageFromBase64(left: double, top: double, base64EncodedImage: String)
Adds an image to the page at the specified position. The image is added to the page's collection of page content objects.

#### Syntax
```js
pageObject.addImageFromBase64(left, top, base64EncodedImage);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|left|double|The left position of the outline.|
|top|double|The top position of the outline.|
|base64EncodedImage|String|A base64-encoded image.|

#### Returns
[Image](image.md)

### addOutline(left: double, top: double, html: String)
Adds an outline to the page at the specified position. The outline is added to the page's collection of page content objects.

#### Syntax
```js
pageObject.addOutline(left, top, html);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|left|double|The left position of the outline.|
|top|double|The top position of the outline.|
|html|String|An HTML string that describes the visual presentation of the outline.|

#### Returns
[Outline](outline.md)

### getContents()
Gets the collection of page content objects for the page.

#### Syntax
```js
pageObject.getContents();
```

#### Parameters
None

#### Returns
[PageContentCollection](pagecontentcollection.md)

### insertPageAsSibling(location: string, title: string)
Inserts a new page as a sibling of the page.

#### Syntax
```js
pageObject.insertPageAsSibling(location, title);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|location|string|The location of the new page.  Possible values are: Before, After|
|title|string|The title of the new page.|

#### Returns
[Page](page.md)

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
