# ConditionalRangeBorderCollection Object (JavaScript API for Excel)

Represents the border objects that make up range border.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|count|int|Number of border objects in the collection. Read-only.|[1.6](../requirement-sets/excel-api-requirement-sets.md)|
|items|[ConditionalRangeBorder[]](conditionalrangeborder.md)|A collection of conditionalRangeBorder objects. Read-only.|[1.6](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|bottom|[ConditionalRangeBorder](conditionalrangeborder.md)|Gets the top border Read-only.|[1.6](../requirement-sets/excel-api-requirement-sets.md)|
|left|[ConditionalRangeBorder](conditionalrangeborder.md)|Gets the top border Read-only.|[1.6](../requirement-sets/excel-api-requirement-sets.md)|
|right|[ConditionalRangeBorder](conditionalrangeborder.md)|Gets the top border Read-only.|[1.6](../requirement-sets/excel-api-requirement-sets.md)|
|top|[ConditionalRangeBorder](conditionalrangeborder.md)|Gets the top border Read-only.|[1.6](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getItem(index: string)](#getitemindex-string)|[ConditionalRangeBorder](conditionalrangeborder.md)|Gets a border object using its name|[1.6](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[ConditionalRangeBorder](conditionalrangeborder.md)|Gets a border object using its index|[1.6](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### getItem(index: string)
Gets a border object using its name

#### Syntax
```js
conditionalRangeBorderCollectionObject.getItem(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|string|Index value of the border object to be retrieved.  Possible values are: EdgeTop, EdgeBottom, EdgeLeft, EdgeRight|

#### Returns
[ConditionalRangeBorder](conditionalrangeborder.md)

### getItemAt(index: number)
Gets a border object using its index

#### Syntax
```js
conditionalRangeBorderCollectionObject.getItemAt(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|Index value of the object to be retrieved. Zero-indexed.|

#### Returns
[ConditionalRangeBorder](conditionalrangeborder.md)
