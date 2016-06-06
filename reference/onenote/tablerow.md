# TableRow Object (JavaScript API for OneNote)

_Applies to: OneNote Online_
_Note: This API is in preview_

Represents a row in a table.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|cellCount|int|Gets the number of cells in the row. Read-only.|
|id|string|Gets the ID of the row. Read-only.|
|rowIndex|int|Gets the index of the row in its parent table. Read-only.|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|cells|[TableCellCollection](tablecellcollection.md)|Gets the cells in the row. Read-only.|
|parentTable|[Table](table.md)|Gets the parent table. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[delete()](#delete)|void|Deletes the entire row.|
|[insertRowAsSibling(insertLocation: string, values: string[])](#insertrowassiblinginsertlocation-string-values-string)|[TableRow](tablerow.md)|Inserts a row before or after the current row.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### delete()
Deletes the entire row.

#### Syntax
```js
tableRowObject.delete();
```

#### Parameters
None

#### Returns
void

### insertRowAsSibling(insertLocation: string, values: string[])
Inserts a row before or after the current row.

#### Syntax
```js
tableRowObject.insertRowAsSibling(insertLocation, values);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|insertLocation|string|Where the new rows should be inserted relative to the current row.  Possible values are: Before, After|
|values|string[]|Optional. Strings to insert in the new row, specified as an array. Must not have more cells than in the current row. Optional.|

#### Returns
[TableRow](tablerow.md)

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
