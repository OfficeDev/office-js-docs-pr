# SectionCollection Object (JavaScript API for OneNote)

_Applies to: OneNote Online_
_Note: This API is in preview_

Represents a collection of sections.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|items|[Section[]](section.md)|A collection of section objects. Read-only.|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getByName(name: string)](#getbynamename-string)|[SectionCollection](sectioncollection.md)|Gets the collection of sections with the specified name.|
|[getItem(index: number or string)](#getitemindex-number-or-string)|[Section](section.md)|Gets a section by ID or by its index in the collection. Read-only.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### getByName(name: string)
Gets the collection of sections with the specified name.

#### Syntax
```js
sectionCollectionObject.getByName(name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|name|string|The name of the section.|

#### Returns
[SectionCollection](sectioncollection.md)

### getItem(index: number or string)
Gets a section by ID or by its index in the collection. Read-only.

#### Syntax
```js
sectionCollectionObject.getItem(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number or string|The ID of the section, or the index location of the section in the collection.|

#### Returns
[Section](section.md)

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
