# InlinePicture object (JavaScript API for Word)

Represents an inline picture.

_Applies to: Word 2016 for Windows_

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|altTextDescription|string|Gets or sets a string that represents the alternative text associated with the inline image|
|altTextTitle|string|Gets or sets a string that contains the title for the inline image.|
|hyperlink|string|Gets or sets the hyperlink associated with the inline image.|
|lockAspectRatio|bool|Gets or sets a value that indicates whether the inline image retains its original proportions when you resize it.|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|height|[float](float.md)|Gets or sets a number that describes the height of the inline image.|
|parentContentControl|[ContentControl](contentcontrol.md)|Gets the content control that contains the inline image. Returns null if there isn't a parent content control. Read-only.|
|width|[float](float.md)|Gets or sets a number that describes the width of the inline image.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getBase64ImageSrc()](#getbase64imagesrc)|string|Gets the base64 encoded string representation of the inline image.|
|[insertContentControl()](#insertcontentcontrol)|[ContentControl](contentcontrol.md)|Wraps the inline picture with a rich text content control.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method details

### getBase64ImageSrc()
Gets the base64 encoded string representation of the inline image.

#### Syntax
```js
inlinePictureObject.getBase64ImageSrc();
```

#### Parameters
None

#### Returns
string

### insertContentControl()
Wraps the inline picture with a rich text content control.

#### Syntax
```js
inlinePictureObject.insertContentControl();
```

#### Parameters
None

#### Returns
[ContentControl](contentcontrol.md)

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


