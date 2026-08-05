---
title: Custom functions and data types
description: Configure Excel custom functions to accept and return formatted numbers, entity values, and other Excel data types.
ms.date: 07/28/2026
ms.topic: overview
ms.custom: scenarios:getting-started
ms.localizationpriority: medium
ai-usage: ai-assisted
---

# Use data types with custom functions in Excel

Excel data types let custom functions accept and return values beyond strings, numbers, Boolean values, and errors. For example, a custom function can return a formatted number or accept an entity value as an argument.

Custom functions and the Excel JavaScript API use the same JSON schema for data types. Once [enabled](#enable-data-types-for-custom-functions), Excel maintains that schema as custom functions calculate and evaluate values.

## Try custom functions with data types

Install [Script Lab](../overview/explore-with-script-lab.md) in Excel, and then run the [Data types: Custom functions](https://github.com/OfficeDev/office-js-snippets/blob/prod/samples/excel/16-custom-functions/data-types-custom-functions.yaml) sample from the **Samples** library.

## How custom functions handle data types

Custom functions can accept data types as parameters and create data types as return values. For an introduction to the available types, see [Overview of data types in Excel add-ins](excel-data-types-overview.md).

> [!NOTE]
> Custom functions don't support the full functionality of the enhanced error objects offered by data types. A custom function can accept a data types error object, but it doesn't maintain the error throughout calculation. At this time, custom functions only support the errors included in the [CustomFunctions.Error object](custom-functions-errors.md).

## Enable data types for custom functions

Custom functions projects include a JSON metadata file, which differs from the JSON schema used by the data types APIs. To enable data types for custom functions, manually add the `allowCustomDataForDataTypeAny` property to the custom functions metadata file and set it to `true`.

```json
"allowCustomDataForDataTypeAny": true
```

For a full description of the manual JSON metadata creation process, see [Manually create JSON metadata for custom functions](custom-functions-json.md). See [allowCustomDataForDataTypeAny](custom-functions-json.md#allowcustomdatafordatatypeany) for additional details about this property.

## Output a formatted number

The following custom function accepts a number and a number format, and then returns a formatted [DoubleCellValue](/javascript/api/excel/excel.doublecellvalue) object.

```js
/**
 * Take a number as the input value and return a double as the output.
 * @customfunction
 * @param {number} value
 * @param {string} format (e.g. "0.00%")
 * @returns A formatted number value.
 */
function createFormattedNumber(value, format) {
    return {
        type: "Double",
        basicValue: value,
        numberFormat: format
    };
}
```

## Input an entity value

The following custom function accepts an [EntityCellValue](/javascript/api/excel/excel.entitycellvalue) object and an attribute name. It returns the entity's `text` property when `attribute` is `text`. Otherwise, it returns the `basicValue` of the requested property.

> [!IMPORTANT]
> When constructing or returning an `EntityCellValue` in a custom function, and that value contains nested entities, only define the `referencedValues` array on the root-level entity. Defining `referencedValues` on a nested entity causes Excel to return a **#VALUE!** error with no JavaScript exception or detailed diagnostics. Use [ReferenceCellValue](/javascript/api/excel/excel.referencecellvalue) indices in nested entities to point to the root entity's `referencedValues` array instead. For more information, see [Entity values](excel-data-types-concepts.md#entity-values).

```js
/**
 * Accept an entity value data type as a function input.
 * @customfunction
 * @param {Excel.EntityCellValue} value
 * @param {string} attribute
 * @returns {any} The text value of the entity.
 */
function getEntityAttribute(value, attribute) {
    if (value.type === "Entity") {
        if (attribute === "text") {
            return value.text;
        } else {
            return value.properties[attribute].basicValue;
        }
    } else {
        return JSON.stringify(value);
    }
}
```

## See also

- [Overview of data types in Excel add-ins](excel-data-types-overview.md)
- [Use data types in Excel add-ins](excel-data-types-concepts.md)
- [Configure your Office Add-in to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
