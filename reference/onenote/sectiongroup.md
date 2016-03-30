# SectionGroup Object (JavaScript API for OneNote)

_Applies to: OneNote Online_
_Note: This API is in preview_


Represents a OneNote section group. Section groups can contain sections and other section groups.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|id|string|Gets the ID of the section group. Read-only.|
|name|string|Gets the name of the section group. Read-only.|


## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|notebook|[Notebook](notebook.md)|Gets the notebook that contains the section group. This value is never null. Read-only.|
|parentSectionGroup|[SectionGroup](sectiongroup.md)|Gets the section group that contains the section group. Returns null if the section group is a direct child of the notebook. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[addSection(title: String)](#addsectiontitle-string)|[Section](section.md)|Adds a new Section to the end of the section group|
|[getSectionGroups()](#getsectiongroups)|[SectionGroupCollection](sectiongroupcollection.md)|Gets the collection of section groups in the section group.|
|[getSections(recursive: bool)](#getsectionsrecursive-bool)|[SectionCollection](sectioncollection.md)|Gets the collection of sections in the section group.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### addSection(title: String)
Adds a new Section to the end of the section group

#### Syntax
```js
sectionGroupObject.addSection(title);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|title|String|Title for the new Section|

#### Returns
[Section](section.md)

### getSectionGroups()
Gets the collection of section groups in the section group.

#### Syntax
```js
sectionGroupObject.getSectionGroups();
```

#### Parameters
None

#### Returns
[SectionGroupCollection](sectiongroupcollection.md)

### getSections(recursive: bool)
Gets the collection of sections in the section group.

#### Syntax
```js
sectionGroupObject.getSections(recursive);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|recursive|bool|true to retrieve all child sections, or false to retrieve immediate child sections only. Default is false.|

#### Returns
[SectionCollection](sectioncollection.md)

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
