# RichText Object (JavaScript API for OneNote)

_Applies to: OneNote Online_
_Note: This API is in preview_


Represents a rich text content object in a paragraph.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|id|string|Gets the ID of the rich text content object. Read-only.|
|text|string|Gets the text of the rich text content object. Read-only.|

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|paragraph|[Paragraph](paragraph.md)|Gets the paragraph that contains the rich text content object. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getHtml()](#gethtml)|[String](string.md)|Gets the HTML of the rich text content object.|
|[insertHtml(html: string, insertLocation: string)](#inserthtmlhtml-string-insertlocation-string)|void|Inserts HTML at the specified location in the rich content object.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### getHtml()
Gets the HTML of the rich text content object.

#### Syntax
```js
richTextObject.getHtml();
```

#### Parameters
None

#### Returns
[String](string.md)

### insertHtml(html: string, insertLocation: string)
Inserts HTML at the specified location in the rich content object.

#### Syntax
```js
richTextObject.insertHtml(html, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|html|string|The HTML to insert.|
|insertLocation|string|The location to insert the HTML.  Possible values are: Before, After|

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
