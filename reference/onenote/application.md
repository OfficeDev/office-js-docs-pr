# Application Object (JavaScript API for OneNote)

_Applies to: OneNote Online_
_Note: This API is in preview_


Represents the top-level object that contains all globally addressable OneNote objects such as notebooks, the active notebook, and the active section.

## Properties

None

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|activeNotebook|[Notebook](notebook.md)|Gets the active notebook. Read-only.|
|notebooks|[NotebookCollection](notebookcollection.md)|Gets the collection of notebooks that are open in the OneNote Application instance. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getNotebookById(id: string)](#getnotebookbyidid-string)|[Notebook](notebook.md)|Gets a notebook by ID.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### getNotebookById(id: string)
Gets a notebook by ID.

#### Syntax
```js
applicationObject.getNotebookById(id);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|id|string|The ID of the notebook.|

#### Returns
[Notebook](notebook.md)

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
