# FilterCriteria Object (JavaScript API for Excel)

Represents the filtering criteria applied to a column.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|color|string|The HTML color string used to filter cells. Used with "cellColor" and "fontColor" filtering.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|criterion1|string|The first criterion used to filter data. Used as an operator in the case of "custom" filtering.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|criterion2|string|The second criterion used to filter data. Only used as an operator in the case of "custom" filtering.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|dynamicCriteria|string|The dynamic criteria from the Excel.DynamicFilterCriteria set to apply on this column. Used with "dynamic" filtering. Possible values are: Unknown, AboveAverage, AllDatesInPeriodApril, AllDatesInPeriodAugust, AllDatesInPeriodDecember, AllDatesInPeriodFebruray, AllDatesInPeriodJanuary, AllDatesInPeriodJuly, AllDatesInPeriodJune, AllDatesInPeriodMarch, AllDatesInPeriodMay, AllDatesInPeriodNovember, AllDatesInPeriodOctober, AllDatesInPeriodQuarter1, AllDatesInPeriodQuarter2, AllDatesInPeriodQuarter3, AllDatesInPeriodQuarter4, AllDatesInPeriodSeptember, BelowAverage, LastMonth, LastQuarter, LastWeek, LastYear, NextMonth, NextQuarter, NextWeek, NextYear, ThisMonth, ThisQuarter, ThisWeek, ThisYear, Today, Tomorrow, YearToDate, Yesterday.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|filterOn|string|The property used by the filter to determine whether the values should stay visible. Possible values are: BottomItems, BottomPercent, CellColor, Dynamic, FontColor, Values, TopItems, TopPercent, Icon, Custom.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|operator|string|The operator used to combine criterion 1 and 2 when using "custom" filtering. Possible values are: And, Or.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|values|object[]|The set of values to be used as part of "values" filtering.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|icon|[Icon](icon.md)|The icon used to filter cells. Used with "icon" filtering.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## Methods
None

