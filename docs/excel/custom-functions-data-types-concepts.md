---
title: Custom functions and custom data types core concepts (preview)
description: 'Learn the core concepts for using Excel custom data types with your custom functions.'
ms.date: 10/22/2021
ms.topic: conceptual
ms.custom: scenarios:getting-started
ms.localizationpriority: medium
---

# Custom functions and custom data types core concepts (preview)

> [!NOTE]
> Custom data types APIs are currently only available in public preview. [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]
> 

Custom data types enhance the Excel JavaScript API by expanding support for data types beyond the original four (string, number, boolean, and error). Custom data types include support for formatted number values, web images, entity values, and improved errors.

## How custom functions handle custom data types

Custom functions can work on top of the data types, can recognize data types and do calculations on data types.

## Scenario + code sample

Scenario:
Take data type as a parameter in a custom function
Custom function calculation results can return data types.

### Formatted number values

Custom functions can return formatted number values as outputs.

## See also

* [Excel custom data types overview](/excel-data-types-overview.md)
* [Excel custom data types core concepts](/excel-data-types-concepts.md)
* [Configure your Office Add-in to use a shared JavaScript runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
