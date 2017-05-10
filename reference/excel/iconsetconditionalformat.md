# IconSetConditionalFormat Object (JavaScript API for Excel)

Represents an IconSet criteria for conditional formatting.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|reverseIconOrder|bool|If true, reverses the icon orders for the IconSet. Note that this cannot be set if custom icons are used.|[1.6](../requirement-sets/excel-api-requirement-sets.md)|
|showIconOnly|bool|If true, hides the values and only shows icons.|[1.6](../requirement-sets/excel-api-requirement-sets.md)|
|style|string|If set, displays the IconSet option for the conditional format. Possible values are: Invalid, ThreeArrows, ThreeArrowsGray, ThreeFlags, ThreeTrafficLights1, ThreeTrafficLights2, ThreeSigns, ThreeSymbols, ThreeSymbols2, FourArrows, FourArrowsGray, FourRedToBlack, FourRating, FourTrafficLights, FiveArrows, FiveArrowsGray, FiveRating, FiveQuarters, ThreeStars, ThreeTriangles, FiveBoxes.|[1.6](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|criteria|[ConditionalIconCriterion](conditionaliconcriterion.md)|An array of Criteria and IconSets for the rules and potential custom icons for conditional icons. Note that for the first criterion only the custom icon can be modified, while type, formula and operator will be ignored when set.|[1.6](../requirement-sets/excel-api-requirement-sets.md)|

## Methods
None

