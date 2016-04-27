# NotebookCollection Object (JavaScript API for OneNote)

_Applies to: OneNote Online_
_Note: This API is in preview_

Represents a collection of notebooks.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|items|[Notebook[]](notebook.md)|A collection of notebook objects. Read-only.|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getByName(name: string)](#getbynamename-string)|[NotebookCollection](notebookcollection.md)|Gets the collection of notebooks with the specified name.|
|[getItem(index: number or string)](#getitemindex-number-or-string)|[Notebook](notebook.md)|Gets a notebook by ID or by its index in the collection. Read-only.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### getByName(name: string)
Gets the collection of notebooks with the specified name.

#### Syntax
```js
notebookCollectionObject.getByName(name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|name|string|The name of the notebook.|

#### Returns
[NotebookCollection](notebookcollection.md)

### getItem(index: number or string)
Gets a notebook by ID or by its index in the collection. Read-only.

#### Syntax
```js
notebookCollectionObject.getItem(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number or string|The ID of the notebook, or the index location of the notebook in the collection.|

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
