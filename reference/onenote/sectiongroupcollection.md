# SectionGroupCollection Object (JavaScript API for OneNote)

_Applies to: OneNote Online_
_Note: This API is in preview_

Represents a collection of section groups.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|items|[SectionGroup[]](sectiongroup.md)|A collection of sectionGroup objects. Read-only.|



## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getByName(name: string)](#getbynamename-string)|[SectionGroupCollection](sectiongroupcollection.md)|Gets the collection of section groups with the specified name.|
|[getItem(index: number or string)](#getitemindex-number-or-string)|[SectionGroup](sectiongroup.md)|Gets a section group by ID or by its index in the collection. Read-only.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### getByName(name: string)
Gets the collection of section groups with the specified name.

#### Syntax
```js
sectionGroupCollectionObject.getByName(name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|name|string|The name of the section group.|

#### Returns
[SectionGroupCollection](sectiongroupcollection.md)

### getItem(index: number or string)
Gets a section group by ID or by its index in the collection. Read-only.

#### Syntax
```js
sectionGroupCollectionObject.getItem(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number or string|The ID of the section group, or the index location of the section group in the collection.|

#### Returns
[SectionGroup](sectiongroup.md)

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
