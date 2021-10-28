---
title: Custom functions and data types overview
description: 'Use Excel data types with your custom functions and Office Add-ins.'
ms.date: 10/27/2021
ms.topic: conceptual
ms.custom: scenarios:getting-started
ms.localizationpriority: high
---

# Use data types with custom functions in Excel (preview)

> [!NOTE]
> The custom functions integration with data types is currently in public preview and is only compatible with Windows devices. To use this feature, you need to [Join the Office Insider Program](https://insider.office.com/) and then choose the **Beta Channel** Insider level. To learn more, see [How to get Office Insider builds on Windows](https://insider.office.com/join/windows).

Data types expand the Excel JavaScript API to support data types beyond the four original data types (string, number, boolean, and error). Data types include support for web images, formatted number values, entity values, and arrays within entity values.

These data types amplify the power of custom functions, because custom functions accept data types as both input and output values. You can generate data types through custom functions, or take existing data types as function arguments into calculations. Once the JSON schema of a data type is set, this schema is maintained throughout custom function calculations.

To learn more about using data types with an Excel add-in, see [Excel data types core concepts](/excel-data-types-concepts.md). To learn more about integrating custom data types with your custom functions, see [Custom functions and data types core concepts](/custom-functions-data-types-concepts.md).

## See also

* [Excel data types overview](/excel-data-types-overview.md)
* [Excel data types core concepts](/excel-data-types-concepts.md)
* [Custom functions and data types core concepts](/custom-functions-data-types-concepts.md)
* [Configure your Office Add-in to use a shared JavaScript runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
