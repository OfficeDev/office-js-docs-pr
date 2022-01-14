---
title: Overview of data types in Excel add-ins
description: 'Data types in the Excel JavaScript API enable Office Add-in developers to work with formatted number values, web images, entity values, arrays within entity values, and enhanced errors as data types.'
ms.date: 12/27/2021
ms.topic: conceptual
ms.prod: excel
ms.custom: scenarios:getting-started
ms.localizationpriority: high
---

# Overview of data types in Excel add-ins (preview)

> [!NOTE]
> Data types APIs are currently only available in public preview. Preview APIs are subject to change and are not intended for use in a production environment. We recommend that you try them out in test and development environments only. Do not use preview APIs in a production environment or within business-critical documents.
>
> To use preview APIs:
>
> - You must reference the **beta** library on the content delivery network (CDN) (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js). The [type definition file](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) for TypeScript compilation and IntelliSense is found at the CDN and [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts). You can install these types with `npm install --save-dev @types/office-js-preview`. For additional information, see the [@microsoft/office-js](https://www.npmjs.com/package/@microsoft/office-js) NPM package readme.
> - You may need to join the [Office Insider program](https://insider.office.com) for access to more recent Office builds.
>
> To try out data types in Office on Windows, you must have an Excel build number greater than or equal to 16.0.14626.10000. To try out data types in Office on Mac, you must have an Excel build number greater than or equal to 16.55.21102600.

Data types in the Excel JavaScript API enable add-in developers to organize complex data structures as objects, such as formatted number values, web images, and entity values.

Prior to the data types addition, the Excel JavaScript API supported string, number, boolean, and error data types. The Excel UI formatting layer is capable of adding currency, date, and other types of formatting to cells that contain the four original data types, but this formatting layer only controls the display of the original data types in the Excel UI. The underlying number value is not changed, even when a cell in the Excel UI is formatted as currency or a date. This gap between an underlying value and the formatted display in the Excel UI can result in confusion and errors during add-in calculations. Custom data types are a solution to this gap.

Data types expand Excel JavaScript API support beyond the four original data types (string, number, boolean, and error) to include web images, formatted number values, entity values, arrays within entity values, and improved error data types as flexible data structures. These types, which power many [linked data types](https://support.microsoft.com/office/what-linked-data-types-are-available-in-excel-6510ab58-52f6-4368-ba0f-6a76c0190772) experiences, allow for precision and simplicity during add-in calculations and extend the potential of Excel add-ins beyond a 2-dimensional grid.

## Data types and custom functions

[!include[Custom functions and data types availability note](../includes/excel-custom-functions-data-types-note.md)]

Data types enhance the power of custom functions. Custom functions accept data types as both inputs to custom functions and outputs of custom functions, and custom functions use the same JSON schema for data types as the Excel JavaScript API. This data types JSON schema is maintained as custom functions calculate and evaluate. To learn more about integrating data types with your custom functions, see [Custom functions and data types](custom-functions-data-types-concepts.md).

## See also

- [Excel data types core concepts](excel-data-types-concepts.md)
- [Excel JavaScript API reference](../reference/overview/excel-add-ins-reference-overview.md)
- [Custom functions and data types](custom-functions-data-types-concepts.md)