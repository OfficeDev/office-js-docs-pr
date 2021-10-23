---
title: Custom functions and custom data types overview (preview)
description: 'Use Excel custom data types with your custom functions and Office Add-ins.'
ms.date: 10/22/2021
ms.topic: conceptual
ms.custom: scenarios:getting-started
ms.localizationpriority: high
---

# Use custom data types with custom functions in Excel (preview)

> [!NOTE]
> Custom data types APIs are currently only available in public preview. [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]
> 

Custom data types expand the Excel JavaScript API to support data types beyond the four original data types (string, number, boolean, and error). Custom data types include support for web images, formatted number values, entity values, and enhanced errors.

These new data types amplify the power of custom functions, because custom functions accept custom data types as both input and output values. You can generate custom data types through custom functions, or take existing custom data types as function arguments into calculations. Once the structure and schema of a custom data type is set, it's maintained throughout add-in and custom function calculations.

To learn more about using custom data types with an Excel add-in, see the [Excel custom data types core concepts](/excel-data-types-concepts.md) article. To learn more about integrating custom data types with your custom functions, see [Custom functions and custom data types core concepts]().

## See also

* [Excel custom data types overview](/excel-data-types-overview.md)
* [Excel custom data types core concepts](/excel-data-types-concepts.md)
* [Configure your Office Add-in to use a shared JavaScript runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
