# BindingSelectionChangedEventArgs Object (JavaScript API for Excel)

Provides information about the binding that raised the SelectionChanged event.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|columnCount|int|Gets the number of columns selected.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|rowCount|int|Gets the number of rows selected.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|startColumn|int|Gets the index of the first column of the selection (zero-based).|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|startRow|int|Gets the index of the first row of the selection (zero-based).|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|binding|[Binding](binding.md)|Gets the Binding object that represents the binding that raised the SelectionChanged event.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## Methods
None

