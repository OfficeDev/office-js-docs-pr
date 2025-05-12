---
title: Create linked entity cell values
description: Create linked entity cell values to represent large data sets in Excel.
ms.topic: how-to
ms.date: 05/12/2025
ms.localizationpriority: medium
---

# Create linked entity cell values

Linked entity cell values integrate data types from external data sources and can display the data as an entity card, like [regular entity values](excel-data-types-entity-card.md). They enable you to scale your data types to represent large data sets without downloading all the data into the workbook. The [Stocks and Geography data domains](https://support.microsoft.com/office/excel-data-types-stocks-and-geography-61a33056-9935-484f-8ac8-f1a89e210877) available via the Excel UI provide linked entity cell values. This article explains how to create your own data provider in an Excel add-in to provide custom values for end users.

Linked entity cell values are linked to an external data source. They provide the following advantages over regular entity values.

- Linked entity cell values can nest, and nested linked entity cell values are not retrieved until referenced; either by the user, or by the worksheet. This helps reduce file size and improve workbook performance.
- Excel uses a cache to allow different cells to reference the same linked entity cell value seamlessly. This also improves workbook performance.

This article expands on information described in the following articles. We recommend reading the following articles before learning how to build your own linked entity cell values.

- [Excel data types: Stocks and geography](https://support.microsoft.com/office/61a33056-9935-484f-8ac8-f1a89e210877)
- [Overview of data types in Excel add-ins](excel-data-types-overview.md)
- [Excel JavaScript API data types entity value card](excel-data-types-entity-card.md)

## Key concepts

Linked entity cell values provide the user with data linked from an external data source. The user can view them as an entity value card.

:::image type="content" source="../images/excel-geography-linked-data-type-seattle.png" alt-text="Screenshot of an entity value data card for the Seattle Geography linked data type in the worksheet.":::

Like regular entity values, linked entity cell values can be referenced in formulas.

:::image type="content" source="../images/excel-geography-seattle-dot-syntax.png" alt-text="Screenshot of using dot notation in a formula using =A1. To display fields for Seattle Geography data type.":::

## Definitions

The following definitions are fundamental to understanding how to implement your own linked entity cell values.

- **Linked entity data domain** – A linked entity data domain describes the overall category that an entity belongs to. Some examples are employees, organizations, or cars.
- **Linked entity cell value** – An instance created from a data domain. An example is an employee value for someone named Joe. It can be displayed as an entity value card.
- **Data provider** - The data provider is recognized by Excel as the source of data for one or more registered linked entity data domains.
- **Linked entity load service function** – Every linked entity data domain defines a load service function to act as the source of data for that domain. The linked entity load service function handles requests from Excel to get linked entity cell values for the workbook. You implement it as a TypeScript or JavaScript custom function.

## How your add-in provides linked entity cell values

This diagram shows the steps that occur when your add-in is loaded and then inserts a new linked entity cell value into a cell. The following description explains what happens at each step of the process.

:::image type="content" source="../images/excel-data-types-linked-entity-workflow.png" alt-text="Diagram showing the five steps for an add-in to register data domains and handle requests from Excel to get properties from a linked entity cell value.":::

1. Excel loads your add-in, and your add-in registers all of the linked entity data domains that it supports. Each registration includes the ID of a linked entity load service function. This ID is called later by Excel to request property values for the linked entity cell value from the linked entity data domain. In this example, one data domain named **Products** is registered.
1. Excel tracks each registered linked entity data domain in a linked entity data domain collection. This enables Excel to call your linked entity load service function when data is needed for a linked entity cell value.
1. Your add-in inserts a new linked entity cell value into the worksheet. In this example a new linked entity cell value is created for the product **Chai**. This would typically occur from the user choosing a button on your add-in that results in creating one or more linked entity cell values. When you create new linked entity cell values, they only contain an initial text string that is displayed in the cell. Excel calls your linked entity load service function to get the remaining property values. Your add-in can also create linked entity cell values from custom functions.
1. Excel calls the linked entity load service function that you registered in step 1. This occurs every time you create a new linked entity cell value, or if a data refresh occurs. Excel calls your linked entity load service function to get all of the property values.
1. The linked entity load service function returns an up-to-date linked entity cell value ([Excel.LinkedEntityCellValue](/javascript/api/excel/excel.linkedentitycellvalue)) for the linked entity ID ([Excel.LinkedEntityId](/javascript/api/excel/excel.linkedentityid)) requested by Excel. Typically, your linked entity load service function queries an external data source to get the values and create the linked entity cell value. In this example the values for **product ID**, **category**, **quantity**, and **price** are returned.

> [!NOTE]
> If Excel needs multiple linked entity cell values, the linked entity IDs are passed as a batch to your linked entity load service function. The linked entity load service then returns a batch result of all values.

The following sections provide additional details about the terms defined earlier in this article.

### Data provider

Your add-in is the data provider and is recognized by Excel as the source of data for one or more registered data domains. Your add-in exposes one or more data provider functions that return data for linked entity cell values. The [data provider](/javascript/api/excel/excel.linkedentitydatadomaincreateoptions) is identified by a text string such as **Contoso** or the name of your add-in. The name must be unique within your add-in.

### Linked entity data domains

The data provider (your add-in) registers one or more data domains. A data domain describes an entity to Excel. For example, a data provider can provide the **products** and **categories** data domains. The domains must be registered with Excel so that it can work with those domains to retrieve and display linked entity cell values and perform calculations.

A data domain describes to Excel the following attributes:

- The name of the data provider it is associated with.
- A domain ID to uniquely identify it, such as **products**.
- A display name for the user, such as **Products**.
- A linked entity load service function to call when Excel needs a linked entity cell value.
- A specified refresh mode and interval describing how often it refreshes.

An example of a linked entity data domain is the **Geography** data domain in Excel that provides linked entity cell values for cities.

### Linked entity cell value

A linked entity cell value is an instance created from a data domain. An example is a value for Seattle, from the [Geography data domain](https://support.microsoft.com/office/excel-data-types-stocks-and-geography-61a33056-9935-484f-8ac8-f1a89e210877). It displays an entity value card like regular entity cell values.

:::image type="content" source="../images/excel-geography-linked-data-type-seattle.png" alt-text="Screenshot of an entity value data card for the Seattle Geography linked data type in the worksheet.":::

Since linked entity cell values are linked to the data domain, they can be refreshed. Also, nested linked entity cell values are not retrieved unless the user requests them (such as viewing the entity card). And nested entity cell values aren’t saved with the worksheet unless they are referenced from the worksheet (such as a formula). This reduces file size and improves performance.

### Linked entity load service function

Each data domain requires a function that Excel can call when it needs linked entity cell values. Your add-in provides the service as a JavaScript or TypeScript function tagged with **@linkedEntityLoadService**. It's recommended to create just one load service function for best performance. Excel sends all requests for linked entity cell values as a batch to the load service function.

## Create a data provider with data domains

The following sections of this article show how to write TypeScript code to implement an Excel add-in that is a data provider for **Contoso**. It provides two data domains named **Products** and **Categories**.

### Register the data domains

Let’s look at the code to register new domains named **Products** and **Categories**. The data provider name is **Contoso**. When the add-in loads, it first registers the data domains with Excel.

Use the [Excel.LinkedEntityDataDomainCreateOptions](/javascript/api/excel/excel.linkedentitydatadomaincreateoptions) type to describe the options you want, including which function to use as the linked entity load service. Then add the domain to the [Workbook.linkedEntityDataDomains](/javascript/api/excel/excel.workbook#excel-excel-workbook-linkedentitydatadomains-member) collection. It's recommended to register domains when you [Initialize your Office Add-in](../develop/initialize-add-in.md).
The following code shows how to register the **Products**, **Categories**, and **Suppliers** data domains.

```typescript
Office.onReady(async () => {
    await Excel.run(async (context) => {
        const productsDomain: Excel.LinkedEntityDataDomainCreateOptions = {
            dataProvider: "Contoso",
            id: "products",
            name: "Products",
            // ID of the custom function that is called on demand by Excel to resolve or refresh linked entity cell values of this data domain.
            loadFunctionId: "CONTOSOLOADSERVICE",
            // periodicRefreshInterval is only required when supportedRefreshModes contains "Periodic".
            periodicRefreshInterval: 300,
            // Manual refresh mode is always supported, even if unspecified.
            supportedRefreshModes: [
                Excel.LinkedEntityDataDomainRefreshMode.periodic,
                Excel.LinkedEntityDataDomainRefreshMode.onLoad
            ]
        };

        const categoriesDomain: Excel.LinkedEntityDataDomainCreateOptions = {
            dataProvider: "Contoso",
            id: "categories",
            name: "Categories",
            loadFunctionId: "CONTOSOLOADSERVICE",
            periodicRefreshInterval: 300,
            supportedRefreshModes: [
                Excel.LinkedEntityDataDomainRefreshMode.periodic,
                Excel.LinkedEntityDataDomainRefreshMode.onLoad
            ]
        };

        const suppliersDomain: Excel.LinkedEntityDataDomainCreateOptions = {
            dataProvider: "Contoso",
            id: "suppliers",
            name: "Suppliers",
            loadFunctionId: "CONTOSOLOADSERVICE"
        };
        // Register the data domains by adding them to the collection.
        context.workbook.linkedEntityDataDomains.add(productsDomain);
        context.workbook.linkedEntityDataDomains.add(categoriesDomain);
        context.workbook.linkedEntityDataDomains.add(suppliersDomain);

        await context.sync();
    });
});
```

## Insert a linked entity cell value

There are two ways to insert a linked entity cell value into a cell on a worksheet.

- Create a command button on the ribbon or a button in your task pane. When the user selects the button, your code inserts a linked entity cell value.
- Create a custom function that returns a linked entity cell value.

The following code example shows how to insert a new linked entity cell value into the selected cell. This code can be called from a command button on the ribbon, or a button in the task pane. Notes about the following code:

- You must specify a `serviceId` of `268436224` for any linked entity cell values you return. This informs Excel that the linked entity cell value is associated with an Excel add-in.
- You must specify a `culture`. Excel will pass it to your linked entity load service function so that you can maintain the original culture when the workbook is opened in a different culture.
- The `text` property is displayed to the user in the cell while the linked entity data value is updated. This prevents the user seeing a blank cell while the update is completed.

```typescript
async function insertProduct() {
    await Excel.run(async (context) => {
        const productLinkedEntity: Excel.LinkedEntityCellValue = {
            type: Excel.CellValueType.linkedEntity,
            id: {
                entityId: "P1", // Don't use exclamation marks in this value.
                domainId: "products", // Don't use exclamation marks in this value.
                serviceId: 268436224,
                culture: "en-US",
            },
            text: "Chai",
        };
        context.workbook.getActiveCell().valuesAsJson = [[productLinkedEntity]];
        await context.sync();
    });
}
```

> [!NOTE]
> Don't use exclamation marks in the `entityID` or `domainId` values.

The following code sample shows how to insert a linked entity cell value by using a custom function. A user could get a linked entity cell value by entering `=CONTOSO.GETPRODUCTBYID("productid")` into any cell. The notes for the previous code sample also apply to this one.

```typescript
/**
 * Custom function that shows how to insert a `LinkedEntityCellValue`.
 * @customfunction
 * @param {string} productID Unique ID of the product.
 * @return {any} `LinkedEntityCellValue` for the requested product, if found.
 */
function getProductById(productID: string): any {
    const product = getProduct(productID);
    if (product === null) {
        throw new CustomFunctions.Error(CustomFunctions.ErrorCode.notAvailable, "Invalid productID");
    }
    const productLinkedEntity: Excel.LinkedEntityCellValue = {
        type: Excel.CellValueType.linkedEntity,
        id: {
            entityId: product.productID,
            domainId: "products",
            serviceId: 268436224,
            culture: "en-US",
        },
        text: product.productName
    };

    return productLinkedEntity;
}
```

## Implement the linked entity load service function

The add-in must provide a linked entity load service function to handle requests from Excel when property values are needed for any linked entity cell values. The function is identified with the `@linkedEntityLoadService` JSDoc tag.

The following code example shows how to create a function that handles data requests from Excel for the **Products** and **Categories** data domains. Notes on the following code:

- It uses helper functions to create the linked entity cell values. That code is shown later.
- If an error occurs it throws a `CustomFunctions.ErrorCode.notAvailable` error. This displays `#CONNECT!` in the cell that the user sees.

```typescript
// Linked entity data domain constants
const productsDomainId = "products";
const categoriesDomainId = "categories";
const suppliersDomainId = "suppliers";

// Linked entity cell value constants
const addinDomainServiceId = 268436224;
const defaultCulture = "en-US";

/**
 * Custom function which acts as the "service" or the data provider for a `LinkedEntityDataDomain`, that is
 * called on demand by Excel to resolve/refresh `LinkedEntityCellValue`s of that `LinkedEntityDataDomain`.
 * @customfunction
 * @linkedEntityLoadService
 * @param {any} request Request to resolve/refresh `LinkedEntityCellValue` objects.
 * @return {any} Resolved/Refreshed `LinkedEntityCellValue` objects that were requested in the passed-in request.
 */
function contosoLoadService(request: any): any {
    const notAvailableError = new CustomFunctions.Error(CustomFunctions.ErrorCode.notAvailable);
    console.log(`Fetching linked entities from request: ${request} ...`);

    try {
        // Parse the request that was passed-in by Excel.
        const parsedRequest: Excel.LinkedEntityLoadServiceRequest = JSON.parse(request);
        // Initialize result to populate and return to Excel.
        const result: Excel.LinkedEntityLoadServiceResult = { entities: [] };

        // Identify the domainId of the request and call the corresponding function to create
        // linked entity cell values for that linked entity data domain.
        for (const { entityId } of parsedRequest.entities) {
            var linkedEntityResult = null;
            switch (parsedRequest.domainId) {
                case productsDomainId: {
                    linkedEntityResult = makeProductLinkedEntity(entityId);
                    break;
                }
                case categoriesDomainId: {
                    linkedEntityResult = makeCategoryLinkedEntity(entityId);
                    break;
                }
                case suppliersDomainId: {
                    linkedEntityResult = makeSupplierLinkedEntity(entityId);
                    break;
                }
                default:
                    throw notAvailableError;
            }

            if (!linkedEntityResult) {
                // Throw an error to signify to Excel that resolution/refresh of the requested linkedEntityId failed.
                throw notAvailableError;
            }

            result.entities.push(linkedEntityResult);
        }

        return result;
    } catch (error) {
        console.error(error);
        throw notAvailableError;
    }
}
```

The following code sample shows the helper function to create a product linked entity cell value. This function is called by the previous code `contosoLoadService` to create a linked entity for a specific product ID. Notes on the following code:

- It uses the same settings as the previous `insertProduct` example for the `type`, `id`, and `text` properties.
- It includes additional properties specific to the **Products** data domain, such as `Product Name` and `Unit Price`.
- It creates a deferred nested linked entity for the category of the product. The properties for the category are not requested until they are needed.

```typescript
/** Helper function to create a linked entity from product properties. */
function makeProductLinkedEntity(productID: string): any {
    // Search the product data in the data source for a matching product ID.
    const product = getProduct(productID);
    if (product === null) {
        // Return null if no matching product is found.
        return null;
    }

    const productLinkedEntity: Excel.LinkedEntityCellValue = {
        type: "LinkedEntity",
        text: product.productName,
        id: {
            entityId: product.productID,
            domainId: productsDomainId,
            serviceId: addinDomainServiceId,
            culture: defaultCulture
        },
        properties: {
            "Product ID": {
                type: "String",
                basicValue: product.productID
            },
            "Product Name": {
                type: "String",
                basicValue: product.productName
            },
            "Quantity Per Unit": {
                type: "String",
                basicValue: product.quantityPerUnit
            },
            // Add Unit Price as a formatted number.
            "Unit Price": {
                type: "FormattedNumber",
                basicValue: product.unitPrice,
                numberFormat: "$* #,##0.00"
            },
            Discontinued: {
                type: "Boolean",
                basicValue: product.discontinued
            }
        },
        layouts: {
            compact: {
                icon: "ShoppingBag"
            },
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
                        collapsed: false,
                        properties: ["Quantity Per Unit", "Unit Price"]
                    },
                    {
                        layout: "List",
                        title: "Additional information",
                        collapsed: true,
                        properties: ["Discontinued"]
                    }
                ]
            }
        }
    };

    // Add image property to the linked entity and then add it to the card layout.
    if (product.productImage) {
        productLinkedEntity.properties["Image"] = {
            type: "WebImage",
            address: product.productImage
        };
        productLinkedEntity.layouts.card.mainImage = { property: "Image" };
    }

    // Add a deferred nested linked entity for the product category.
    const category = getCategory(product.categoryID.toString());
    if (category) {
        productLinkedEntity.properties["Category"] = {
            type: "LinkedEntity",
            text: category.categoryName,
            id: {
                entityId: category.categoryID.toString(),
                domainId: categoriesDomainId,
                serviceId: addinDomainServiceId,
                culture: defaultCulture
            }
        };

        // Add nested product category to the card layout.
        productLinkedEntity.layouts.card.sections[0].properties.push("Category");
    }

    // Add a deferred nested linked entity for the supplier.
    const supplier = getSupplier(product.supplierID.toString());
    if (supplier) {
        productLinkedEntity.properties["Supplier"] = {
            type: "LinkedEntity",
            text: supplier.companyName,
            id: {
                entityId: supplier.supplierID.toString(),
                domainId: suppliersDomainId,
                serviceId: addinDomainServiceId,
                culture: defaultCulture
            }
        };

        // Add nested product supplier to the card layout.
        productLinkedEntity.layouts.card.sections[2].properties.push("Supplier");
    }

    return productLinkedEntity;
}
```

The following code sample shows the helper function to create a category linked entity cell value. This function is called by the previous code `contosoLoadService` to create a linked entity for a specific category ID.

```typescript
/** Helper function to create a linked entity from category properties. */
function makeCategoryLinkedEntity(categoryID: string): any {
    // Search the sample JSON category data for a matching category ID.
    const category = getCategory(categoryID);
    if (category === null) {
        // Return null if no matching category is found.
        return null;
    }

    const categoryLinkedEntity: Excel.LinkedEntityCellValue = {
        type: "LinkedEntity",
        text: category.categoryName,
        id: {
            entityId: category.categoryID,
            domainId: categoriesDomainId,
            serviceId: addinDomainServiceId,
            culture: defaultCulture
        },
        properties: {
            "Category ID": {
                type: "String",
                basicValue: category.categoryID,
                propertyMetadata: {
                    // Exclude the category ID property from the card view and auto-complete.
                    excludeFrom: {
                        cardView: true,
                        autoComplete: true
                    }
                }
            },
            "Category Name": {
                type: "String",
                basicValue: category.categoryName
            },
            Description: {
                type: "String",
                basicValue: category.description
            }
        },
        layouts: {
            compact: {
                icon: "Branch"
            }
        }
    };

    return categoryLinkedEntity;
}
```

The following code sample shows the helper function to create a supplier linked entity cell value. This function is called by the previous code `contosoLoadService` to create a linked entity for a specific supplier ID.

```typescript
/** Helper function to create linked entity from supplier properties. */
function makeSupplierLinkedEntity(supplierID: string): any {
    // Search the sample JSON category data for a matching supplier ID.
    const supplier = getSupplier(supplierID);
    if (supplier === null) {
        // Return null if no matching supplier is found.
        return null;
    }

    const supplierLinkedEntity: Excel.LinkedEntityCellValue = {
        type: "LinkedEntity",
        text: supplier.companyName,
        id: {
            entityId: supplier.supplierID,
            domainId: suppliersDomainId,
            serviceId: addinDomainServiceId,
            culture: defaultCulture
        },
        properties: {
            "Supplier ID": {
                type: "String",
                basicValue: supplier.supplierID
            },
            "Company Name": {
                type: "String",
                basicValue: supplier.companyName
            },
            "Contact Name": {
                type: "String",
                basicValue: supplier.contactName
            },
            "Contact Title": {
                type: "String",
                basicValue: supplier.contactTitle
            }
        },
        cardLayout: {
            title: { property: "Company Name" },
            sections: [
                {
                    layout: "List",
                    properties: ["Supplier ID", "Company Name", "Contact Name", "Contact Title"]
                }
            ]
        }
    };

    return supplierLinkedEntity;
}
```

The following code sample contains sample data you can use with the previous code samples.

```typescript
/// Sample product data.
const products = [
    {
        productID: "P1",
        productName: "Chai",
        supplierID: "S1",
        categoryID: "C1",
        quantityPerUnit: "10 boxes x 20 bags",
        unitPrice: 18,
        discontinued: false,
        productImage: "https://upload.wikimedia.org/wikipedia/commons/thumb/0/04/Masala_Chai.JPG/320px-Masala_Chai.JPG"
    }
];

/// Sample product category data.
const categories = [
    {
        categoryID: "C1",
        categoryName: "Beverages",
        description: "Soft drinks, coffees, teas, beers, and ales"
    }];

/// Sample product supplier data.
const suppliers = [
    {
        supplierID: "S1",
        companyName: "Exotic Liquids",
        contactName: "Ema Vargova",
        contactTitle: "Purchasing Manager"
    }];
```

## Data refresh options

When you register a data domain, the user can refresh it manually at any time, such as by choosing **Refresh All** from the **Data** tab. There are three refresh modes you can specify for your data domain.

- `manual`- The data is refreshed only when the user chooses to refresh. This is the default mode. Manual refresh can always be performed by the user, even when the refresh mode is set to `onLoad` or `periodic`.
- `onLoad`- The data is refreshed when the data domain is registered (typically when the add-in is loaded). Afterwards, data is only refreshed manually by the user. If you want to refresh data when the workbook is opened, configure your add-in to load on document open. For more information, see [Run code in your Office Add-in when the document opens](../develop/run-code-on-document-open.md).
- `periodic`-  The data is refreshed when the data domain is registered (typically when the add-in is loaded). Afterwards, the data is continuously updated after a specified interval of time. For example you could specify that the data domain refreshes every 300 seconds (which is the minimum value). The number of seconds is always rounded up to the nearest number of minutes since the refresh interval is only performed in minutes.

The following code example shows how to configure a data domain to refresh on load, and then continue to refresh every 5 minutes.

```typescript
const productsDomain: Excel.LinkedEntityDataDomainCreateOptions = {
    dataProvider: domainDataProvider,
    id: "products",
    name: "Products",
    // ID of the custom function that is called on demand by Excel to resolve or refresh linked entity cell values of this data domain.
    loadFunctionId: loadFunctionId,
    // periodicRefreshInterval is only required when supportedRefreshModes contains "Periodic".
    periodicRefreshInterval: 300, // equivalent to 5 minutes.
    // Manual refresh mode is always supported, even if unspecified.
    supportedRefreshModes: [
        Excel.LinkedEntityDataDomainRefreshMode.periodic,
        Excel.LinkedEntityDataDomainRefreshMode.onLoad
    ]
};
```

You can also programmatically request a refresh on a linked entity data domain by using either of the following methods.

- `LinkedEntityDataDomain.refresh()` - Refreshes all `LinkedEntityCellValue` objects of the linked entity data domain.
- `LinkedEntityDataDomainCollection.refreshAll()` - Refreshes all `LinkedEntityCellValue` objects of all linked entity data domains in the collection.

The refresh methods request a refresh which occurs asynchronously. To determine the results of the refresh, listen for the `onRefreshCompleted` event. The following code sample shows an example of listening for the `onRefreshCompleted` event.

```typescript
await Excel.run(async (context) => {
    const dataDomains = context.workbook.linkedEntityDataDomains;
    dataDomains.onRefreshCompleted.add(onLinkedEntityDomainRefreshed);

    await context.sync();
});

async function onLinkedEntityDomainRefreshed(eventArgs: Excel.LinkedEntityDataDomainRefreshCompletedEventArgs): Promise<any> {
    console.log("Linked entity domain refreshed: " + eventArgs.id);
    console.log("Refresh status: " + eventArgs.refreshed);
    console.log("Refresh error: " + eventArgs.errors);
    return null;
}
```

## Error handling with the linked entity load service

When Excel calls your add-in to get data for a linked entity cell value, it's possible an error can occur. If Excel is unable to connect to your add-in at all, such as when the add-in isn't loaded, Excel displays the `#CONNECT!` error to the user.

If your linked entity load service function encounters an error, it should throw the `notAvailableError` error. This causes Excel to show `#CONNECT!` to the user.

The following code shows how to handle an error in a linked entity load service function.

```typescript
async function contosoLoadService(request: any): Promise<any> {
    const notAvailableError = new CustomFunctions.Error(CustomFunctions.ErrorCode.notAvailable);
    try {
        // Create and return a new linked entity cell value.
        let linkedEntityResult = ...
      ...
        if (!linkedEntityResult) {
            // Throw an error to signify to Excel that resolution or refresh of the requested linkedEntityId failed.
            throw notAvailableError;
        }
      ...
    } catch (error) {
        console.error(error);
        throw notAvailableError;
    }
}
```

## Debugging the linked entity load service

Most add-in functionality for linked entity data types can be debugged using the guidance in [Overview of debugging Office Add-ins](../testing/debug-add-ins-overview.md). However, the linked entity load service function can be implemented in a shared runtime or a JavaScript-only runtime (also know as a custom functions runtime.) If you choose to implement the function in a JavaScript-only runtime, use the [Custom functions debugging in a non-shared runtime](custom-functions-debugging.md) guidance.

The linked entity load service function uses the custom functions architecture, regardless of which runtime you use. However, there are significant differences from regular custom functions.

Linked entity load service functions have the following differences from custom functions:

- They don't appear to end users for usage in formulas.
- They don't support the JSDoc tags `@streaming` or `@volatile`. The user will see a **#CALC!** error if these tags are used.

Linked entity load service functions have the following similarities with custom functions:

- They use [Custom functions naming and localization](custom-functions-naming.md).
- They use the same error handling approach.

## Behavior in Excel 2019 and earlier

If someone opens a worksheet with linked entity cell values on an older version of Excel that doesn’t support linked entity cell values, Excel shows the cell values as errors. This is the designed behavior. This is also why you set the `basicType` to `Error` and the `basicValue` to `#VALUE!` every time you insert or update a linked entity cell value. This is the error that Excel uses as a fallback on older versions.

## Best practices

- Don't use exclamation marks in the `entityID` or `domainId` values.
- Register linked entity data domains in the initialization code `Office.OnReady` so that the user has immediate functionality such as the ability to refresh the linked entity cell values.
- After publishing your add-in, don’t change the linked entity data domain IDs. Consistent IDs across the same logical objects helps with performance.
- Always provide the `text` property when creating a new linked entity cell value. This value is displayed while Excel calls your data provider function to get the remaining property values. Otherwise the user sees a blank cell until the data is retrieved.

## See also

- [Overview of data types in Excel add-ins](excel-data-types-overview.md)
