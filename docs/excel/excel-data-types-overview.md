---
title: Overview of data types in Excel add-ins
description: Data types in the Excel JavaScript API enable Office Add-in developers to work with formatted number values, web images, entities, arrays within entities, and enhanced errors as data types.
ms.date: 10/10/2022
ms.topic: conceptual
ms.prod: excel
ms.custom: scenarios:getting-started
ms.localizationpriority: high
---

# Overview of data types in Excel add-ins

Data types organize complex data structures as objects. This includes formatted number values, web images, and entities as [entity cards](excel-data-types-entity-card.md).

Prior to the data types addition, the Excel JavaScript API supported string, number, boolean, and error data types. The Excel UI formatting layer is capable of adding currency, date, and other types of formatting to cells that contain the four original data types, but this formatting layer only controls the display of the original data types in the Excel UI. The underlying number value is not changed, even when a cell in the Excel UI is formatted as currency or a date. This gap between an underlying value and the formatted display in the Excel UI can result in confusion and errors during add-in calculations. The data types APIs are a solution to this gap.

Data types expand Excel JavaScript API support beyond the four original data types (string, number, boolean, and error) to include [web images](excel-data-types-concepts.md#web-image-values), [formatted number values](excel-data-types-concepts.md#formatted-number-values), [entities](excel-data-types-concepts.md#entity-values), arrays within entities, and improved [error data types](excel-data-types-concepts.md#improved-error-support) as flexible data structures. These types, which power many [linked data types](https://support.microsoft.com/office/what-linked-data-types-are-available-in-excel-6510ab58-52f6-4368-ba0f-6a76c0190772) experiences, allow for precision and simplicity during add-in calculations and extend the potential of Excel add-ins beyond a 2-dimensional grid.

To learn how to use data types APIs, start with the [Excel data types core concepts](excel-data-types-concepts.md) article.

## Data types and custom functions

Data types enhance the power of custom functions. Custom functions accept data types as both inputs to custom functions and outputs of custom functions, and custom functions use the same JSON schema for data types as the Excel JavaScript API. This data types JSON schema is maintained as custom functions calculate and evaluate. To learn more about integrating data types with your custom functions, see [Custom functions and data types](custom-functions-data-types-concepts.md).

## See also

- [Excel data types core concepts](excel-data-types-concepts.md)
- [Use cards with entity value data types](excel-data-types-entity-card.md)
- [Excel JavaScript API reference](../reference/overview/excel-add-ins-reference-overview.md)
- [Custom functions and data types](custom-functions-data-types-concepts.md)