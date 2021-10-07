---
title: Excel JavaScript API custom data types overview
description: 'Custom data types in the Excel JavaScript API enable Office Add-in developers to work with formatted number values, rich errors, and web images as data types.'
ms.date: 10/06/2021
ms.topic: conceptual
ms.prod: excel
ms.custom: scenarios:getting-started
ms.localizationpriority: high
---

# Create custom data types with Excel add-ins (preview)

> [!NOTE]
> Custom data types APIs are currently only available in public preview. [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]
> 

Custom data types in the Excel JavaScript API enable add-in developers to work with complex data types such as formatted number values, rich errors, web images, and entities.

Prior to the custom data types addition, the Excel JavaScript API supported string, number, boolean, and error data types. The Excel UI formatting layer is capable of adding currency, date, and other types of formatting to cell that contain the four original data types, but this formatting layer only controls the display of the original data types in the Excel UI. The underlying number value is not changed, even when a cell is formatted as currency or a date. This gap between an underlying value and the formatted display in the Excel UI can result in confusion and errors during add-in calculations.

Custom data types expand Excel JavaScript API support beyond the four original data types (string, number, boolean, and error) to include web images, formatted number values, entities, and rich error data types. These custom data types allow for precision and simplicity during calculations.

```js
Code sample
```

## See also

* [Excel custom data types core concepts](/excel-data-types-concepts.md)
* [Excel JavaScript API reference](../reference/overview/excel-add-ins-reference-overview.md)