# PivotTableCollection Object (JavaScript API for Excel)

Represents a collection of all the PivotTables that are part of the workbook or worksheet.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[PivotTable[]](pivottable.md)|A collection of pivotTable objects. Read-only.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getCount()](#getcount)|int|Gets the number of pivot tables in the collection.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(name: string)](#getitemname-string)|[PivotTable](pivottable.md)|Gets a PivotTable by name.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNullObject(name: string)](#getitemornullobjectname-string)|[PivotTable](pivottable.md)|Gets a PivotTable by name. If the PivotTable does not exist, will return a null object.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[refreshAll()](#refreshall)|void|Refreshes all the pivot tables in the collection.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### getCount()
Gets the number of pivot tables in the collection.

#### Syntax
```js
pivotTableCollectionObject.getCount();
```

#### Parameters
None

#### Returns
int

### getItem(name: string)
Gets a PivotTable by name.

#### Syntax
```js
pivotTableCollectionObject.getItem(name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|name|string|Name of the PivotTable to be retrieved.|

#### Returns
[PivotTable](pivottable.md)

### getItemOrNullObject(name: string)
Gets a PivotTable by name. If the PivotTable does not exist, will return a null object.

#### Syntax
```js
pivotTableCollectionObject.getItemOrNullObject(name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|name|string|Name of the PivotTable to be retrieved.|

#### Returns
[PivotTable](pivottable.md)

### refreshAll()
Refreshes all the pivot tables in the collection.

#### Syntax
```js
pivotTableCollectionObject.refreshAll();
```

#### Parameters
None

#### Returns
void
