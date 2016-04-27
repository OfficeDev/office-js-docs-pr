# Notebook Object (JavaScript API for OneNote)

_Applies to: OneNote Online_
_Note: This API is in preview_

Represents a OneNote notebook. Notebooks contain section groups and sections.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|id|string|Gets the ID of the notebook. Read-only.|
|name|string|Gets the name of the notebook. Read-only.|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[addSection(title: String)](#addsectiontitle-string)|[Section](section.md)|Adds a new section to the end of the notebook.|
|[getSectionGroups()](#getsectiongroups)|[SectionGroupCollection](sectiongroupcollection.md)|Gets the section groups in the notebook.|
|[getSections(recursive: bool)](#getsectionsrecursive-bool)|[SectionCollection](sectioncollection.md)|Gets the sections in the notebook.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### addSection(title: String)
Adds a new section to the end of the notebook.

#### Syntax
```js
notebookObject.addSection(title);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|title|String|The name of the new section.|

#### Returns
[Section](section.md)

### getSectionGroups()
Gets the section groups in the notebook.

#### Syntax
```js
notebookObject.getSectionGroups();
```

#### Parameters
None

#### Returns
[SectionGroupCollection](sectiongroupcollection.md)

### getSections(recursive: bool)
Gets the sections in the notebook.

#### Syntax
```js
notebookObject.getSections(recursive);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|recursive|bool|true to retrieve all child sections, or false to retrieve direct child sections only. Default is false.|

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
