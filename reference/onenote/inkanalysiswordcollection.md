# InkAnalysisWordCollection Object (JavaScript API for OneNote)

_Applies to: OneNote Online_  
_Note: This API is in preview_  


Represents a collection of InkAnalysisWord objects.

## Properties

| Property	   | Type	|Description|Feedback|
|:---------------|:--------|:----------|:-------|
|count|int|Returns the number of InkAnalysisWords in the page. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWordCollection-count)|
|items|[InkAnalysisWord[]](inkanalysisword.md)|A collection of inkAnalysisWord objects. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWordCollection-items)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Feedback|
|:---------------|:--------|:----------|:-------|
|[getItem(index: number or string)](#getitemindex-number-or-string)|[InkAnalysisWord](inkanalysisword.md)|Gets a InkAnalysisWord object by ID or by its index in the collection. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWordCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[InkAnalysisWord](inkanalysisword.md)|Gets a InkAnalysisWord on its position in the collection.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWordCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWordCollection-load)|

## Method Details


### getItem(index: number or string)
Gets a InkAnalysisWord object by ID or by its index in the collection. Read-only.

#### Syntax
```js
inkAnalysisWordCollectionObject.getItem(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number or string|The ID of the InkAnalysisWord object, or the index location of the InkAnalysisWord object in the collection.|

#### Returns
[InkAnalysisWord](inkanalysisword.md)

### getItemAt(index: number)
Gets a InkAnalysisWord on its position in the collection.

#### Syntax
```js
inkAnalysisWordCollectionObject.getItemAt(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|Index value of the object to be retrieved. Zero-indexed.|

#### Returns
[InkAnalysisWord](inkanalysisword.md)

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
