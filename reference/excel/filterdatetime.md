# FilterDatetime Object (JavaScript API for Excel)

Represents how to filter a date when filtering on values.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|date|string|The date in ISO8601 format used to filter data.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|specificity|string|How specific the date should be used to keep data. For example, if the date is 2005-04-02 and the specifity is set to "month", the filter operation will keep all rows with a date in the month of april 2009. Possible values are: Year, Monday, Day, Hour, Minute, Second.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods
None

