# PivotTable Object (JavaScript API for Excel)

Represents an Excel PivotTable.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|id|string|Id of the PivotTable. Read-only.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|Name of the PivotTable.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|worksheet|[Worksheet](worksheet.md)|The worksheet containing the current PivotTable. Read-only.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[refresh()](#refresh)|void|Refreshes the PivotTable.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### refresh()
Refreshes the PivotTable.

#### Syntax
```js
pivotTableObject.refresh();
```

#### Parameters
None

#### Returns
void
