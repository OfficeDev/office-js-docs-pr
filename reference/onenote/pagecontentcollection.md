# PageContentCollection Object (JavaScript API for OneNote)

_Applies to: OneNote Online_
_Note: This API is in preview_


Represents the contents of a page, as a collection of page content objects.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|items|[PageContent[]](pagecontent.md)|A collection of pageContent objects. Read-only.|


## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getItem(index: number or string)](#getitemindex-number-or-string)|[PageContent](pagecontent.md)|Gets a page content object by its index in the collection. Read-only.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### getItem(index: number or string)
Gets a page content object by its index in the collection. Read-only.

#### Syntax
```js
pageContentCollectionObject.getItem(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number or string|A number or ID that identifies the index location of the page content object.|

#### Returns
[PageContent](pagecontent.md)

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
