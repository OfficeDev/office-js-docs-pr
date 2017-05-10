# ConditionalRangeFormat Object (JavaScript API for Excel)

A format object encapsulating the conditional formats range's font, fill, borders, and other properties.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|numberFormat|object|Represents Excel's number format code for the given range. Cleared if null is passed in.|[1.6](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|borders|[ConditionalRangeBorderCollection](conditionalrangebordercollection.md)|Collection of border objects that apply to the overall conditional format range. Read-only.|[1.6](../requirement-sets/excel-api-requirement-sets.md)|
|fill|[ConditionalRangeFill](conditionalrangefill.md)|Returns the fill object defined on the overall conditional format range. Read-only.|[1.6](../requirement-sets/excel-api-requirement-sets.md)|
|font|[ConditionalRangeFont](conditionalrangefont.md)|Returns the font object defined on the overall conditional format range. Read-only.|[1.6](../requirement-sets/excel-api-requirement-sets.md)|

## Methods
None

