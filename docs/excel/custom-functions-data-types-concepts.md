---
title: Custom functions and data types core concepts
description: 'Learn the core concepts for using Excel data types with your custom functions.'
ms.date: 11/01/2021
ms.topic: conceptual
ms.custom: scenarios:getting-started
ms.localizationpriority: medium
---

# Custom functions and data types core concepts (preview)

[!include[Custom functions and data types availability note](../includes/excel-custom-functions-data-types-note.md)]

Data types enhance the Excel JavaScript API by expanding support for data types beyond the original four (string, number, boolean, and error). Data types include support for formatted number values, web images, entity values, and arrays within entity values. Custom functions accept data types as both input and output values, expanding the calculation power of custom functions.

To learn more about using data types with an Excel add-in, see [Excel data types core concepts](/excel-data-types-concepts.md).

## How custom functions handle data types

Custom functions can recognize data types and accept them as parameter values. A custom function can create a new data type for a return value. Custom functions use the same JSON schema for data types as the Excel JavaScript API, and this JSON schema is maintained as custom functions calculate and evaluate.

> [!NOTE]
> Custom functions do not support the full functionality of the enhanced error objects offered by data types. A custom function can accept a data types error object, but it won't be maintained throughout calculation. At this time, custom functions only support the errors included in the [CustomFunctions.Error object](/custom-functions-errors.md).

## Enable data types for custom functions

To use this feature, you need to manually update your JSON metadata. For more temporary testing, you can customize your Script Lab settings instead of manually updating JSON metadata. The following sections outline these steps in more detail.

### Manually update JSON metadata

Custom functions projects include a JSON metadata file. This JSON metadata file differs from the JSON schema used by data types APIs. To use the data types integration with custom functions, the custom functions JSON metadata file must be manually updated to include the property `allowCustomDataForDataTypeAny`. Set this property to `true`.

For a full description of the manual JSON creation process, see [Manually create JSON metadata for custom functions](custom-functions-json.md). See [allowCustomDataForDataTypeAny](custom-functions-json.md#allowcustomdatafordatatypeany-preview) for additional details about this property.

### Script Lab option

The custom functions integration with data types is available for testing with Script Lab, in addition to the manual JSON metadata update described in the preceding section. To learn more about Script Lab, see [Explore Office JavaScript API using Script Lab](../overview/explore-with-script-lab.md). To test this feature with Script Lab, update the settings using the following steps.

1. Open the Script Lab **Code** task pane.
1. In the lower right corner, select the **Settings** button.
1. Go to the **User Settings** tab and enter `allowCustomDataForDataTypeAny: true`.

![Screenshot showing the steps to enable data types for custom functions in Script Lab.](../images/custom-functions-script-lab-data-type.png)

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

* [Custom functions and data types overview](/custom-functions-data-types-overview.md)
* [Overview of data types in Excel add-ins](/excel-data-types-overview.md)
* [Excel data types core concepts](/excel-data-types-concepts.md)
* [Configure your Office Add-in to use a shared JavaScript runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
