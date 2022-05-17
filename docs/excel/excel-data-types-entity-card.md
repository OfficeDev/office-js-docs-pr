---
title: Excel JavaScript API data types entity value card
description: 
ms.date: 05/17/2022
ms.topic: conceptual
ms.prod: excel
ms.localizationpriority: medium
---

# Excel data types entity value card (preview)

> [!NOTE]
> Data types APIs are currently only available in public preview. Preview APIs are subject to change and are not intended for use in a production environment. We recommend that you try them out in test and development environments only. Do not use preview APIs in a production environment or within business-critical documents.
>
> To use preview APIs:
>
> - You must reference the **beta** library on the content delivery network (CDN) (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js). The [type definition file](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) for TypeScript compilation and IntelliSense is found at the CDN and [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts). You can install these types with `npm install --save-dev @types/office-js-preview`. For additional information, see the [@microsoft/office-js](https://www.npmjs.com/package/@microsoft/office-js) NPM package readme.
> - You may need to join the [Office Insider program](https://insider.office.com) for access to more recent Office builds.
>
> To try out data types in Office on Windows, you must have an Excel build number greater than or equal to 16.0.14626.10000. To try out data types in Office on Mac, you must have an Excel build number greater than or equal to 16.55.21102600.

This article describes how to use the [Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) to work with the card component of data types entity values. An entity value is a container for data types, similar to an object in object oriented programming. The card component is an optional pop up window for an entity data type, displaying additional information about the entity value in a cell. This article introduces properties, layout options for the card, and data attribution functionality.

## Properties

The entity value [`properties`](/javascript/api/excel/excel.entitycellvalue#excel-excel-entitycellvalue-properties-member) property allows you to set customized information about your set of data types. The following code sample shows the JSON for an entity value with multiple custom properties. The `properties` field accepts `string` values, and each string value must have a `type` and `basicValue` setting.

```json
const entity: Excel.EntityCellValue = {
    type: Excel.CellValueType.entity,
    text: productName,
    properties: {
        "Product ID": {
            type: Excel.CellValueType.string,
            basicValue: productID.toString() || ""
        },
        "Product Name": {
            type: Excel.CellValueType.string,
            basicValue: productName || ""
        },
        "Quantity Per Unit": {
            type: Excel.CellValueType.string,
            basicValue: product.quantityPerUnit || ""
        },
        // Set Unit Price as a formatted number.
        "Unit Price": {
            type: Excel.CellValueType.formattedNumber,
            basicValue: product.unitPrice,
            numberFormat: "$* #,##0.00"
        },
        Discontinued: {
            type: Excel.CellValueType.boolean,
            basicValue: product.discontinued || false
        }
    },
    layouts: {
        // Enter layout settings here.
    }
};
```

## Card layout

The entity value [`layouts`](/javascript/api/excel/excel.entitycellvalue#excel-excel-entitycellvalue-layouts-member) property allows the option to create a [`card`](/javascript/api/excel/excel.entityviewlayouts) for the entity. The `card` field offers properties of the [CardLayoutStandardProperties](/javascript/api/excel/excel.cardlayoutstandardproperties) object: `title`, `subTitle`, `sections`, and `mainImage`.

The following JSON code sample shows a `card` layout with a `title` setting and three `sections` within the card. The `sections` field uses the [`CardLayoutSectionStandardProperties`](javascript/api/excel/excel.cardlayoutsectionstandardproperties) object. 

Each card section can have a `layout`, `title`, and `properties`. The `layout` field uses the [`CardLayoutListSection`](/javascript/api/excel/excel.cardlayoutlistsection) object and accepts the property `list`. The `title` field accepts `string` values, and the `properties` field accepts arrays of strings. Sections can also be collapsible, and can be defined as collapsed or not collapsed when the entity card is opened in the Excel UI.

```json
const entity: Excel.EntityCellValue = {
    type: Excel.CellValueType.entity,
    text: productName,
    properties: {
        // Enter property settings here.
    },
    layouts: {
        card: {
            title: { property: "Product Name" },
            sections: [
                {
                    layout: "List",
                    properties: ["Product ID"]
                },
                {
                    layout: "List",
                    title: "Quantity and price",
                    collapsible: true,
                    collapsed: false, // This section will not be collapsed when the card is opened.
                    properties: ["Quantity Per Unit", "Unit Price"]
                },
                {
                    layout: "List",
                    title: "Additional information",
                    collapsed: true, // This section will be collapsed when the card is opened.
                    properties: ["Discontinued"]
                }
            ]
        }
    }
};
```

## Data attribution

The `provider` property offers the `description`, `logoSourceAddress`, and `logoTargetAddress` fields.

```json
const entity: Excel.EntityCellValue = {
    type: Excel.CellValueType.entity,
    text: productName,
    properties: {
        // Enter property settings here.
    },
    layouts: {
        // Enter layout settings here.
    },
    provider: {
        description: product.providerName, // Name of the data provider. Displays as a tooltip when hovering over the logo. Also displays as a fallback if the source address for the image is broken.
        logoSourceAddress: product.sourceAddress, // Source URL of the logo to display.
        logoTargetAddress: product.targetAddress // Destination URL that the logo navigates to when clicked.
    }
};
```

## See also

- [Overview of data types in Excel add-ins](excel-data-types-overview.md)
- [Excel data types core concepts](excel-data-types-concepts.md)
- [Excel JavaScript API reference](../reference/overview/excel-add-ins-reference-overview.md)
