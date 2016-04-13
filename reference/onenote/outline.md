# Outline Object (JavaScript API for OneNote)

_Applies to: OneNote Online_
_Note: This API is in preview_

Represents a container for Paragraph objects.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|id|string|Gets the ID of the Outline object. Read-only.|



## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|pageContent|[PageContent](pagecontent.md)|Gets the PageContent object that contains the Outline. This object defines the position of the Outline on the page. Read-only.|
|paragraphs|[ParagraphCollection](paragraphcollection.md)|Gets the collection of Paragraph objects in the Outline. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[append(html: string)](#appendhtml-string)|void|Adds the specified HTML to the bottom of the Outline.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|
|[prepend(html: string)](#prependhtml-string)|void|Adds the specified HTML to the top of the Outline.|

## Method Details


### append(html: string)
Adds the specified HTML to the bottom of the Outline.

#### Syntax
```js
outlineObject.append(html);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|html|string|HTML string to append.|

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

### prepend(html: string)
Adds the specified HTML to the top of the Outline.

#### Syntax
```js
outlineObject.prepend(html);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|html|string|HTML string to prepend.|

#### Returns
void
