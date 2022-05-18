---
title: Excel JavaScript API data types entity value card
description: Learn how to use entity value cards with data types in your Excel add-in.
ms.date: 05/18/2022
ms.topic: conceptual
ms.prod: excel
ms.localizationpriority: medium
---

# Excel data types entity value cards (preview)

> [!NOTE]
> Data types APIs are currently only available in public preview. Preview APIs are subject to change and are not intended for use in a production environment. We recommend that you try them out in test and development environments only. Do not use preview APIs in a production environment or within business-critical documents.
>
> To use preview APIs:
>
> - You must reference the **beta** library on the content delivery network (CDN) (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js). The [type definition file](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) for TypeScript compilation and IntelliSense is found at the CDN and [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts). You can install these types with `npm install --save-dev @types/office-js-preview`. For additional information, see the [@microsoft/office-js](https://www.npmjs.com/package/@microsoft/office-js) NPM package readme.
> - You may need to join the [Office Insider program](https://insider.office.com) for access to more recent Office builds.
>
> To try out data types in Office on Windows, you must have an Excel build number greater than or equal to 16.0.14626.10000. To try out data types in Office on Mac, you must have an Excel build number greater than or equal to 16.55.21102600.

This article describes how to use the [Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) to work with the card component of data types entity values. An entity value, or [EntityCellValue](/javascript/api/excel/excel.entitycellvalue), is a container for data types, similar to an object in object oriented programming. The card component is a pop up window for an entity data type that displays additional information about the entity value in a cell. This article introduces card properties, layout options for the card, and card data attribution functionality.

The following screenshot shows a list of grocery store products and an open entity value card for the **Tofu** product from the list.

:::image type="content" source="../images/excel-data-types-entity-card-tofu.png" alt-text="A screenshot showing an entity value data type with the card window displayed.":::

## Card properties

The entity value [`properties`](/javascript/api/excel/excel.entitycellvalue#excel-excel-entitycellvalue-properties-member) property allows you to set customized information about your data types. The `properties` key accepts nested data types. Each nested property, or data type, must have a `type` and `basicValue` setting.

> [!IMPORTANT]
> The nested `properties` data types are used in combination with the [Card layout](#card-layout) values described in the subsequent article section. After defining a nested data type in `properties`, it must be assigned in the `layouts` property to display on the card.

The following code snippet shows the JSON for an entity value with multiple custom properties.

> [!NOTE]
> The following code snippet is an excerpt. To see the complete code sample, visit the [OfficeDev/office-js-snippets](https://github.com/OfficeDev/office-js-snippets/blob/main/samples/excel/85-preview-apis/data-types-entity-values.yaml) repository.

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

The following screenshot shows an entity value card that uses the preceding code snippet. The screenshot shows the **Product ID**, **Product Name**, **Quantity Per Unit**, and **Unit Price** information from the preceding code snippet.

:::image type="content" source="../images/excel-data-types-entity-card-properties.png" alt-text="A screenshot showing an entity value data type with the card layout window displayed. The card shows the product name, product ID, quantity per unit, and unit price information.":::

## Card layout

The entity value [`layouts`](/javascript/api/excel/excel.entitycellvalue#excel-excel-entitycellvalue-layouts-member) property creates a [`card`](/javascript/api/excel/excel.entityviewlayouts) for the entity and then specifies the appearance of that card, such as the title of the card, an image for the card, and the number of sections to display.

> [!IMPORTANT]
> The nested `layouts` values are used in combination with the [Card properties](#card-properties) data types described in the preceding article section. A nested data type must be defined in `properties` before it can be assigned in `layouts` to display on the card.

Nested within the `card` property, use the [`CardLayoutStandardProperties`](/javascript/api/excel/excel.cardlayoutstandardproperties) object. The `CardLayoutStandardProperties` object offers the `title`, `subTitle`, `sections`, and `mainImage` properties.

The following entity value JSON code snippet shows a `card` layout with a nested `title` object and three `sections` within the card. Note that the `title` property `"Product Name"` has a corresponding data type in the preceding article section. The `sections` property takes a nested array and uses the [`CardLayoutSectionStandardProperties`](/javascript/api/excel/excel.cardlayoutsectionstandardproperties) object to define the appearance of each section.

Within each card section you can specify `layout`, `title`, and `properties`. The `layout` key uses the [`CardLayoutListSection`](/javascript/api/excel/excel.cardlayoutlistsection) object and accepts the value `"List"`. The `title` key of the `sections` property accepts `string` values, and the `properties` key accepts an array of strings. Sections can also be collapsible and can be defined with boolean values as collapsed or not collapsed when the entity card is opened in the Excel UI.

> [!NOTE]
> The following code snippet is an excerpt. To see the complete code sample, visit the [OfficeDev/office-js-snippets](https://github.com/OfficeDev/office-js-snippets/blob/main/samples/excel/85-preview-apis/data-types-entity-values.yaml) repository.

```json
const entity: Excel.EntityCellValue = {
    type: Excel.CellValueType.entity,
    text: productName,
    properties: {
        // Enter property settings here.
    },
    layouts: {
        card: {
            title: { 
                property: "Product Name" 
            },
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
                    collapsible: true,
                    collapsed: true, // This section will be collapsed when the card is opened.
                    properties: ["Discontinued"]
                }
            ]
        }
    }
};
```

The following screenshot shows an entity value card that uses the preceding code snippet. The screenshot shows the `title` object, which uses the **Product Name** and is set to **Pavlova**. The screenshot also shows `sections`. The **Quantity and price** section is collapsible and contains **Quantity Per Unit** and **Unit Price**. The **Additional information** field is collapsible and is collapsed when the card is opened.

:::image type="content" source="../images/excel-data-types-entity-card-sections.png" alt-text="A screenshot showing an entity value data type with the card layout window displayed. The card shows the card title and sections.":::

## Card data attribution

Entity value cards can display a data attribution to give credit to the provider of the information in the entity card. The entity value [`provider`](/javascript/api/excel/excel.entitycellvalue#excel-excel-entitycellvalue-provider-member) property uses the [`CellValueProviderAttributes`](/javascript/api/excel/excel.cellvalueproviderattributes) object, which  defines the `description`, `logoSourceAddress`, and `logoTargetAddress` values.

The data provider property displays an image in the lower left corner of the entity card, using the `logoSourceAddress` to specify the image. The `logoTargetAddress` value defines the URL destination if the logo image is selected or clicked. The `description` value displays as a tooltip when hovering over the logo. The `description` value also displays as a plain text fallback if the `logoSourceAddress` is not defined or if the source address for the image is broken.

The following JSON code snippet shows an entity value that uses the `provider` property to specify a data provider attribution for the entity.

> [!NOTE]
> The following code snippet is an excerpt. To see the complete code sample, visit the [OfficeDev/office-js-snippets](https://github.com/OfficeDev/office-js-snippets/blob/main/samples/excel/85-preview-apis/data-types-entity-attribution.yaml) repository.

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

The following screenshot shows an entity value card that uses the preceding code snippet. The screenshot shows the data provider attribution in the lower left corner. In this instance, the data provider is Microsoft and the Microsoft logo is displayed.

:::image type="content" source="../images/excel-data-types-entity-card-attribution.png" alt-text="A screenshot showing an entity value data type with the card layout window displayed. The card shows the data provider attribution in the lower left corner.":::

## See also

- [Overview of data types in Excel add-ins](excel-data-types-overview.md)
- [Excel data types core concepts](excel-data-types-concepts.md)
- [Excel JavaScript API reference](../reference/overview/excel-add-ins-reference-overview.md)