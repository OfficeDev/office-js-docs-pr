# RangeView Object (JavaScript API for Excel)

RangeView represents a set of visible cells of the parent range.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|cellAddresses|object[][]|Represents the cell addresses of the RangeView. Read-only.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|columnCount|int|Returns the number of visible columns. Read-only.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|formulas|object[][]|Represents the formula in A1-style notation.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|formulasLocal|object[][]|Represents the formula in A1-style notation, in the user's language and number-formatting locale.  For example, the English "=SUM(A1, 1.5)" formula would become "=SUMME(A1; 1,5)" in German.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|formulasR1C1|object[][]|Represents the formula in R1C1-style notation.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|index|int|Returns a value that represents the index of the RangeView. Read-only.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|numberFormat|object[][]|Represents Excel's number format code for the given cell.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|rowCount|int|Returns the number of visible rows. Read-only.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|text|object[][]|Text values of the specified range. The Text value will not depend on the cell width. The # sign substitution that happens in Excel UI will not affect the text value returned by the API. Read-only.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|valueTypes|string|Represents the type of data of each cell. Read-only. Possible values are: Unknown, Empty, String, Integer, Double, Boolean, Error.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|values|object[][]|Represents the raw values of the specified range view. The data returned could be of type string, number, or a boolean. Cell that contain an error will return the error string.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|rows|[RangeViewCollection](rangeviewcollection.md)|Represents a collection of range views associated with the range. Read-only.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getRange()](#getrange)|[Range](range.md)|Gets the parent range associated with the current RangeView.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### getRange()
Gets the parent range associated with the current RangeView.

#### Syntax
```js
rangeViewObject.getRange();
```

#### Parameters
None

#### Returns
[Range](range.md)
