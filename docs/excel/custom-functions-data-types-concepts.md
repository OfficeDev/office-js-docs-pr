---
title: Custom functions and custom data types core concepts (preview)
description: 'Learn the core concepts for using Excel custom data types with your custom functions.'
ms.date: 10/23/2021
ms.topic: conceptual
ms.custom: scenarios:getting-started
ms.localizationpriority: medium
---

# Custom functions and custom data types core concepts (preview)

> [!NOTE]
> Custom data types APIs are currently only available in public preview. [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]
> 

Custom data types enhance the Excel JavaScript API by expanding support for data types beyond the original four (string, number, boolean, and error). Custom data types include support for formatted number values, web images, and entity values. Custom functions accept custom data types as both input and output values, expanding the calculation power of custom functions.

To learn more about using custom data types with an Excel add-in, see the [Excel custom data types core concepts (preview)](/excel-data-types-concepts.md) article

## How custom functions handle custom data types

Custom functions can recognize custom data types and take these data types as both input and output values. A custom function can generate a new custom data type or accept an existing custom data type as an input parameter. Custom functions use the same JSON schema for custom data types as the Excel JavaScript API, and this JSON schema is maintained as custom functions calculate and evaluate.

> [!NOTE]
> Custom functions do not support the full functionality of the enhanced error objects offered by custom data types. Custom functions can accept custom data types error objects, but the custom data type error object may not be maintained throughout calculation. At this time, custom functions support the errors included in the [CustomFunctions.Error object](/custom-functions-errors.md).

## How to enable custom data types for custom functions

The custom data types integration with custom functions is currently only available in public preview. To try out this new feature, you need to join the Office Insider program, adjust your Office update channel, and customize your Script Lab settings. The following sections outline these steps in more detail.

### Office Insider program

To try out the custom functions integration with custom data types, you first need to [join Office Insider](https://insider.office.com/join). After joining the Office Insider program, set your **Office update channel** to the "Beta Channel".

### Script Lab settings

Script Lab requires a settings update to enable the custom data types integration with custom functions. Take the following steps to update the Script Lab settings.

1. Open the Script Lab **Code** task pane.
1. In the lower right corner, select the **Settings** button.
1. In the lower left corner, change the environment from **Production** to **Alpha**.
1. Go to the **User Settings** tab and enter `allowCustomDataForDataTypeAny: true`.

![Screenshot showing the steps to enable custom data types for custom functions in Script Lab.](../images/custom-functions-script-lab-data-type.png)

## Scenario and code sample: Formatted number values

[Section in progress]
The following code sample shows how to take a formatted number value as an input parameter for a custom function. The following custom function returns a formatted number value as an output.

```js
// Sample in progress.
/**
 * Create Entity from sales data.
 * @customfunction
 * @param {any[][]} data Description here.
 * @returns{any[][]} Description here.
 */
function createEntity(data) {
    yearEntity["properties"] = {
        Inventory: {
            type: "FormattedNumber",
            basicValue: item.inventories,
            numberFormat: "[Red][>=100];[Blue]"
        },
        Profit: {
            type: "FormattedNumber",
            basicValue: profit,
            numberFormat: '_($* #,##0.00_);_($* (#,##0.00);_($* " -"??_);_(@_)'
        }
    };
}
```

## See also

* [Custom functions and custom data types overview (preview)](/custom-functions-data-types-overview.md)
* [Excel custom data types overview](/excel-data-types-overview.md)
* [Excel custom data types core concepts](/excel-data-types-concepts.md)
* [Configure your Office Add-in to use a shared JavaScript runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
