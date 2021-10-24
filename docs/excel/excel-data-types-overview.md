---
title: Excel JavaScript API custom data types overview
description: 'Custom data types in the Excel JavaScript API enable Office Add-in developers to work with formatted number values, rich errors, and web images as data types.'
ms.date: 10/08/2021
ms.topic: conceptual
ms.prod: excel
ms.custom: scenarios:getting-started
ms.localizationpriority: high
---

# Create custom data types with Excel add-ins (preview)

> [!NOTE]
> Custom data types APIs are currently only available in public preview. [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]
> 

Custom data types in the Excel JavaScript API enable add-in developers to organize complex data structures as objects, such as formatted number values, web images, rich errors, and entities.

Prior to the custom data types addition, the Excel JavaScript API supported string, number, boolean, and error data types. The Excel UI formatting layer is capable of adding currency, date, and other types of formatting to cell that contain the four original data types, but this formatting layer only controls the display of the original data types in the Excel UI. The underlying number value is not changed, even when a cell in the Excel UI is formatted as currency or a date. This gap between an underlying value and the formatted display in the Excel UI can result in confusion and errors during add-in calculations.

Custom data types expand Excel JavaScript API support beyond the four original data types (string, number, boolean, and error) to include web images, formatted number values, entities, and rich error data types as flexible data structures. These custom data types allow for precision and simplicity during add-in calculations and extend the power of Excel add-ins beyond a 2-dimensional grid.

```js
// Scenario
```

## Custom data types and custom functions

Custom data types enhance the power of custom functions. Custom functions accept custom data types as both inputs to custom functions and outputs of custom functions. To learn more about integrating custom data types with your custom functions, see [Custom functions and custom data types core concepts](/custom-functions-data-types-concepts.md).

## See also

* [Excel custom data types core concepts](/excel-data-types-concepts.md)
* [Excel JavaScript API reference](../reference/overview/excel-add-ins-reference-overview.md)
* [Custom functions and custom data types overview](/custom-functions-data-types-overview.md)