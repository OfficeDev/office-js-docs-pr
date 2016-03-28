# SortField object (JavaScript API for Excel)

_Applies to: Excel 2016, Excel Online, Excel for iOS, Office 2016_

Represents a condition in a sorting operation.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|ascending|bool|Represents whether the sorting is done in an ascending fashion.|
|color|string|Represents the color that is the target of the condition if the sorting is on font or cell color.|
|dataOption|string|Represents additional sorting options for this field. Possible values are: Normal, TextAsNumber.|
|key|int|Represents the column (or row, depending on the sort orientation) that the condition is on. Represented as an offset from the first column (or row).|
|sortOn|string|Represents the type of sorting of this condition. Possible values are: Value, CellColor, FontColor, Icon.|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|icon|[Icon](icon.md)|Represents the icon that is the target of the condition if the sorting is on the cell's icon.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in the JavaScript layer, with property and object values specified in the parameter.|

## Method Details


### load(param: object)
Fills the proxy object created in the JavaScript layer, with property and object values specified in the parameter.

#### Syntax
```js
object.load(param);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|param|object|Optional. Accepts parameter and relationship names as a delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns
void
