---
title: Custom functions and data types
description: Use Excel data types with your custom functions and Office Add-ins.
ms.date: 10/13/2022
ms.topic: conceptual
ms.custom: scenarios:getting-started
ms.localizationpriority: medium
---

# Use data types with custom functions in Excel

Data types expand the Excel JavaScript API to support data types beyond the original four cell value types (string, number, boolean, and error). Data types include support for web images, formatted number values, entities, and arrays within entities.

These data types amplify the power of custom functions, because custom functions accept data types as both input and output values. You can generate data types through custom functions, or take existing data types as function arguments into calculations. Once the JSON schema of a data type is set, this schema is maintained throughout the calculations.

To learn more about using data types with an Excel add-in, see [Overview of data types in Excel add-ins](excel-data-types-overview.md).

## How custom functions handle data types

Custom functions can recognize data types and accept them as parameter values. A custom function can create a new data type for a return value. Custom functions use the same JSON schema for data types as the Excel JavaScript API, and this JSON schema is maintained as custom functions calculate and evaluate.

> [!NOTE]
> Custom functions do not support the full functionality of the enhanced error objects offered by data types. A custom function can accept a data types error object, but it won't be maintained throughout calculation. At this time, custom functions only support the errors included in the [CustomFunctions.Error object](custom-functions-errors.md).

## Enable data types for custom functions

Custom functions projects include a JSON metadata file. This JSON metadata file differs from the JSON schema used by data types APIs. To use the data types integration with custom functions, the custom functions JSON metadata file must be manually updated to include the property `allowCustomDataForDataTypeAny`. Set this property to `true`.

For a full description of the manual JSON metadata creation process, see [Manually create JSON metadata for custom functions](custom-functions-json.md). See [allowCustomDataForDataTypeAny](custom-functions-json.md#allowcustomdatafordatatypeany) for additional details about this property.

## Output a formatted number value

The following code sample shows how to create a [FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue) data type with a custom function. The function takes a basic number and a format setting as the input parameters and returns a formatted number value data type as the output.

```js
/**
 * Take a number as the input value and return a formatted number value as the output.
 * @customfunction
 * @param {number} value
 * @param {string} format (e.g. "0.00%")
 * @returns A formatted number value.
 */
function createFormattedNumber(value, format) {
    return {
        type: "FormattedNumber",
        basicValue: value,
        numberFormat: format
    }
}
```

## Input an entity value

The following code sample shows a custom function that takes an [EntityCellValue](/javascript/api/excel/excel.entitycellvalue) data type as an input. If the `attribute` parameter is set to `text`, then the function returns the `text` property of the entity value. Otherwise, the function returns the `basicValue` property of the entity value.

```js
/**
 * Accept an entity value data type as a function input.
 * @customfunction
 * @param {any} value
 * @param {string} attribute
 * @returns {any} The text value of the entity.
 */
function getEntityAttribute(value, attribute) {
    if (value.type == "Entity") {
        if (attribute == "text") {
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

* [Overview of data types in Excel add-ins](excel-data-types-overview.md)
* [Excel data types core concepts](excel-data-types-concepts.md)
* [Configure your Office Add-in to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
