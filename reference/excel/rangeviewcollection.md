# RangeViewCollection Object (JavaScript API for Excel)

Represents a collection of RangeView objects.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[RangeView[]](rangeview.md)|A collection of rangeView objects. Read-only.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getCount()](#getcount)|int|Gets the number of RangeView objects in the collection.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[RangeView](rangeview.md)|Gets a RangeView Row via it's index. Zero-Indexed.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### getCount()
Gets the number of RangeView objects in the collection.

#### Syntax
```js
rangeViewCollectionObject.getCount();
```

#### Parameters
None

#### Returns
int

### getItemAt(index: number)
Gets a RangeView Row via it's index. Zero-Indexed.

#### Syntax
```js
rangeViewCollectionObject.getItemAt(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|Index of the visible row.|

#### Returns
[RangeView](rangeview.md)
