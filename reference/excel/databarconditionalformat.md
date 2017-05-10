# DataBarConditionalFormat Object (JavaScript API for Excel)

Represents an Excel Conditional Data Bar Type.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|axisColor|string|HTML color code representing the color of the Axis line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|[1.6](../requirement-sets/excel-api-requirement-sets.md)|
|axisFormat|string|Representation of how the axis is determined for an Excel data bar. Possible values are: Automatic, None, CellMidPoint.|[1.6](../requirement-sets/excel-api-requirement-sets.md)|
|barDirection|string|Represents the direction that the data bar graphic should be based on. Possible values are: Context, LeftToRight, RightToLeft.|[1.6](../requirement-sets/excel-api-requirement-sets.md)|
|showDataBarOnly|bool|If true, hides the values from the cells where the data bar is applied.|[1.6](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|lowerBoundRule|[ConditionalDataBarRule](conditionaldatabarrule.md)|The rule for what consistutes the lower bound (and how to calculate it, if applicable) for a data bar.|[1.6](../requirement-sets/excel-api-requirement-sets.md)|
|negativeFormat|[ConditionalDataBarNegativeFormat](conditionaldatabarnegativeformat.md)|Representation of all values to the left of the axis in an Excel data bar. Read-only.|[1.6](../requirement-sets/excel-api-requirement-sets.md)|
|positiveFormat|[ConditionalDataBarPositiveFormat](conditionaldatabarpositiveformat.md)|Representation of all values to the right of the axis in an Excel data bar. Read-only.|[1.6](../requirement-sets/excel-api-requirement-sets.md)|
|upperBoundRule|[ConditionalDataBarRule](conditionaldatabarrule.md)|The rule for what constitutes the upper bound (and how to calculate it, if applicable) for a data bar.|[1.6](../requirement-sets/excel-api-requirement-sets.md)|

## Methods
None

