# Outline Object (JavaScript API for OneNote)

_Applies to: OneNote Online_
_Note: This API is in preview_


Represents a region on a page that contains paragraphs. An outline can be positioned on the page.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|id|string|Gets the ID of the outline. Read-only.|


## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|pageContent|[PageContent](pagecontent.md)|Gets the page content object that contains the outline. Read-only.|
|paragraphs|[ParagraphCollection](paragraphcollection.md)|Gets the paragraphs in the outline. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[append(html: string)](#appendhtml-string)|void|Appends the specified HTML to the outline.|
|[getHtml()](#gethtml)|string|Gets the HTML representation of the outline.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|
|[prepend(html: string)](#prependhtml-string)|void|Prepend the specified HTML to the outline.|

## Method Details


### append(html: string)
Appends the specified HTML to the outline.

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

### getHtml()
Gets the HTML representation of the outline.

#### Syntax
```js
outlineObject.getHtml();
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

### prepend(html: string)
Prepend the specified HTML to the outline.

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
