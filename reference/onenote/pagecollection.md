# PageCollection Object (JavaScript API for OneNote)

_Applies to: OneNote Online_
_Note: This API is in preview_


Represents a collection of pages.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|items|[Page[]](page.md)|A collection of page objects. Read-only.|

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getByTitle(title: string)](#getbytitletitle-string)|[PageCollection](pagecollection.md)|Gets the collection of pages with the specified title.|
|[getItem(index: number or string)](#getitemindex-number-or-string)|[Page](page.md)|Gets a page by its index in the collection. Read-only.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### getByTitle(title: string)
Gets the collection of pages with the specified title.

#### Syntax
```js
pageCollectionObject.getByTitle(title);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|title|string|The title of the page.|

#### Returns
[PageCollection](pagecollection.md)

### getItem(index: number or string)
Gets a page by its index in the collection. Read-only.

#### Syntax
```js
pageCollectionObject.getItem(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number or string|A number or ID that identifies the index location of the page.|

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
