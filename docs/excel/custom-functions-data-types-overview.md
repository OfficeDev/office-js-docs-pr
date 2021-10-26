---
title: Custom functions and custom data types overview (preview)
description: 'Use Excel custom data types with your custom functions and Office Add-ins.'
ms.date: 10/25/2021
ms.topic: conceptual
ms.custom: scenarios:getting-started
ms.localizationpriority: high
---

# Use custom data types with custom functions in Excel (preview)

> [!NOTE]
> The custom functions integration with custom data types is currently only available in public preview and is only compatible with Windows devices. To use this feature, you need to join the [Office Insider program](https://insider.office.com/) and then choose the **Beta Channel** Insider level. See [Join the Office Insider Program](https://insider.office.com/join/windows) to learn more.

Custom data types expand the Excel JavaScript API to support data types beyond the four original data types (string, number, boolean, and error). Custom data types include support for web images, formatted number values, and entity values.

These custom data types amplify the power of custom functions, because custom functions accept custom data types as both input and output values. You can generate custom data types through custom functions, or take existing custom data types as function arguments into calculations. Once the JSON schema of a custom data type is set, this schema is maintained throughout custom function calculations.

To learn more about using custom data types with an Excel add-in, see the [Excel custom data types core concepts](/excel-data-types-concepts.md) article. To learn more about integrating custom data types with your custom functions, see [Custom functions and custom data types core concepts](/custom-functions-data-types-concepts.md).

## See also

* [Excel custom data types overview](/excel-data-types-overview.md)
* [Excel custom data types core concepts](/excel-data-types-concepts.md)
* [Custom functions and custom data types core concepts](/custom-functions-data-types-concepts.md)
* [Configure your Office Add-in to use a shared JavaScript runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
