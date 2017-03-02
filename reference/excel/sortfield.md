# SortField Object (JavaScript API for Excel)

Represents a condition in a sorting operation.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|ascending|bool|Represents whether the sorting is done in an ascending fashion.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|color|string|Represents the color that is the target of the condition if the sorting is on font or cell color.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|dataOption|string|Represents additional sorting options for this field. Possible values are: Normal, TextAsNumber.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|key|int|Represents the column (or row, depending on the sort orientation) that the condition is on. Represented as an offset from the first column (or row).|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|sortOn|string|Represents the type of sorting of this condition. Possible values are: Value, CellColor, FontColor, Icon.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|icon|[Icon](icon.md)|Represents the icon that is the target of the condition if the sorting is on the cell's icon.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## Methods
None

