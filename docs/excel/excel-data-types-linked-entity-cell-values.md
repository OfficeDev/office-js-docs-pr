---
title: Create linked entity data types in Excel add-ins
description: Learn how to register linked entity data domains, insert linked entity cell values, and refresh external data in Excel add-ins.
ai-usage: ai-assisted
ms.topic: how-to
ms.date: 06/30/2026
ms.localizationpriority: medium
---

# Create linked entity data types in Excel add-ins

Use linked entity cell values when your Excel add-in needs to show external data without storing the full dataset in the workbook. For example, you can show a product, customer, or supplier in one cell, open a rich card for details, and load nested data only when Excel needs it.

Excel uses the same pattern for built-in linked data types such as [Stocks and Geography](https://support.microsoft.com/office/excel-data-types-stocks-and-geography-61a33056-9935-484f-8ac8-f1a89e210877). This article shows how to build your own data provider in an Excel add-in.

In this article, you'll learn how to:

- Register linked entity data domains.
- Insert linked entity cell values from a command or custom function.
- Implement a linked entity load service function.
- Configure refresh and error handling.

Before you start, review these related articles:

- [Overview of data types in Excel add-ins](excel-data-types-overview.md)
- [Use data types in Excel add-ins](excel-data-types-concepts.md)
- [Use cards with cell value data types](excel-data-types-entity-card.md)
- [Add properties to Excel basic cell values](excel-data-types-add-properties-to-basic-cell-values.md)

## What linked entity cell values do

Linked entity cell values connect workbook cells to an external data source and display the result as an entity card.

:::image type="content" source="../images/excel-geography-linked-data-type-seattle.png" alt-text="Screenshot of an entity value data card for the Seattle Geography linked data type in the worksheet.":::

Like regular entity values, you can reference linked entity cell values in formulas.

:::image type="content" source="../images/excel-geography-seattle-dot-syntax.png" alt-text="Screenshot of using dot notation in a formula using =A1. To display fields for Seattle Geography data type.":::

Compared with regular entity values, linked entity cell values offer these advantages.

- Nested linked entity cell values aren't retrieved until the user or worksheet references them. This helps reduce file size and improve workbook performance.
- Excel caches linked entity cell values so multiple cells can reference the same value efficiently.

## Key terms

Use the following terms throughout this article.

- **Linked entity data domain** – A linked entity data domain describes the overall category that an entity belongs to. Some examples are employees, organizations, or cars.
- **Linked entity cell value** – An instance created from a data domain. An example is an employee value for someone named Joe. It can be displayed as an entity value card.
- **Data provider** - The data provider is recognized by Excel as the source of data for one or more registered linked entity data domains.
- **Linked entity load service function** – Every linked entity data domain defines a load service function to act as the source of data for that domain. The linked entity load service function handles requests from Excel to get linked entity cell values for the workbook. You implement it as a TypeScript or JavaScript custom function.

## How the linked entity flow works

This diagram shows what happens after your add-in loads and inserts a linked entity cell value into a cell.

:::image type="content" source="../images/excel-data-types-linked-entity-workflow.png" alt-text="Diagram showing the five steps for an add-in to register data domains and handle requests from Excel to get properties from a linked entity cell value.":::

1. Excel loads your add-in, and your add-in registers all of the linked entity data domains that it supports. Each registration includes the ID of a linked entity load service function. Excel calls this ID later to request property values for the linked entity cell value from the linked entity data domain. In this example, one data domain named **Products** is registered.
1. Excel tracks each registered linked entity data domain in a linked entity data domain collection. This enables Excel to call your linked entity load service function when data is needed for a linked entity cell value.
1. Your add-in inserts a new linked entity cell value into the worksheet. In this example, you create a new linked entity cell value for the product **Chai**. This step typically occurs when the user chooses a button on your add-in that results in creating one or more linked entity cell values. When you create new linked entity cell values, they only contain an initial text string that is displayed in the cell. Excel calls your linked entity load service function to get the remaining property values. Your add-in can also create linked entity cell values from custom functions.
1. Excel calls the linked entity load service function that you registered in step 1. This occurs every time you create a new linked entity cell value, or if a data refresh occurs. Excel calls your linked entity load service function to get all of the property values.
1. The linked entity load service function returns an up-to-date linked entity cell value ([Excel.LinkedEntityCellValue](/javascript/api/excel/excel.linkedentitycellvalue)) for the linked entity ID ([Excel.LinkedEntityId](/javascript/api/excel/excel.linkedentityid)) requested by Excel. Typically, your linked entity load service function queries an external data source to get the values and create the linked entity cell value. In this example, the values for **product ID**, **category**, **quantity**, and **price** are returned.

> [!NOTE]
> If Excel needs multiple linked entity cell values, the linked entity IDs are passed as a batch to your linked entity load service function. The linked entity load service then returns a batch result of all values.

The following sections provide additional details about the terms defined earlier in this article.

### Data provider

Your add-in is the data provider and is recognized by Excel as the source of data for one or more registered data domains. Your add-in exposes one or more data provider functions that return data for linked entity cell values. In [Excel.LinkedEntityDataDomainCreateOptions](/javascript/api/excel/excel.linkedentitydatadomaincreateoptions), set `dataProvider` to a text string such as **Contoso** or the name of your add-in. The name must be unique within your add-in.

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

Since linked entity cell values are linked to the data domain, they can be refreshed. When you implement nested linked entity cell values, note the following behaviors that reduce file size to improve performance.

- Nested linked entity cell values aren't retrieved unless the user requests them, such as by viewing the entity card.
- Nested linked entity cell values aren't saved with the worksheet unless the worksheet references them, such as a formula.

### Linked entity load service function

Each data domain requires a function that Excel can call when it needs linked entity cell values. Your add-in provides the service as a JavaScript or TypeScript function tagged with **@linkedEntityLoadService**. It's recommended to create just one load service function for best performance. Excel sends all requests for linked entity cell values as a batch to the load service function.

## Create a data provider with data domains

The following sections show how to write TypeScript code for an Excel add-in that acts as the data provider for **Contoso**. In this example, the add-in provides three data domains: **Products**, **Categories**, and **Suppliers**.

### Register the data domains

Start by registering each data domain that your add-in supports. In this example, the data provider name is **Contoso** and the domains are **Products**, **Categories**, and **Suppliers**.

Use [Excel.LinkedEntityDataDomainCreateOptions](/javascript/api/excel/excel.linkedentitydatadomaincreateoptions) to define each domain, including the linked entity load service function that Excel should call. Then add each domain to the [Workbook.linkedEntityDataDomains](/javascript/api/excel/excel.workbook#excel-excel-workbook-linkedentitydatadomains-member) collection. Register domains when you [initialize your Office Add-in](../develop/initialize-add-in.md).

The following code registers the **Products**, **Categories**, and **Suppliers** data domains.

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

There are two ways to insert a linked entity cell value into a worksheet cell.

- Create a command button on the ribbon or a button in the task pane. When the user selects the button, your code inserts a linked entity cell value.
- Create a custom function that returns a linked entity cell value.

The following example inserts a new linked entity cell value into the selected cell. You can call this code from a ribbon command or from a button in the task pane.

Keep the following requirements in mind:

- You must specify a `serviceId` of `268436224` for any linked entity cell values you return. This informs Excel that the linked entity cell value is associated with an Excel add-in.
- You must specify a `culture`. Excel passes this value to your linked entity load service function so that you can maintain the original culture when the workbook is opened in a different culture.
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

The add-in must provide a linked entity load service function to handle requests from Excel when property values are needed for any linked entity cell values. The function is identified with the `@linkedEntityLoadService` JSDoc tag. Depending on the version of the Excel client in which the add-in is installed, to load the linked entity cell values, implement one of the following options in your linked entity load service function.

- Create helper functions to create the `LinkedEntityCellValue` objects.
- Register an event handler for the [onLinkedEntityCellValueLoaded](/javascript/api/excel/excel.linkedentitydatadomaincollection#excel-excel-linkedentitydatadomaincollection-onlinkedentitycellvalueloaded-member) event. Then, call [loadLinkedEntityCellValue](/javascript/api/excel/excel.linkedentitydatadomaincollection#excel-excel-linkedentitydatadomaincollection-loadlinkedentitycellvalue-member(1)), passing the linked entity ID as a parameter. The `loadLinkedEntityCellValue` API is supported in ExcelApi 1.21 and later.

The following code examples show how to create a function that uses each of these implementations to handle data requests from Excel for the **Products**, **Categories**, and **Suppliers** data domains. Select the tab for your preferred implementation option.

# [Helper function](#tab/helper-function)

### Create helper functions

The following code sample shows the helper function to create a product linked entity cell value. This function is called by the `contosoLoadService` load service function to create a linked entity for a specific product ID. Notes on the following code.

- It uses the same settings as the previous `insertProduct` example for the `type`, `id`, and `text` properties.
- It includes additional properties specific to the **Products** data domain, such as `Product Name` and `Unit Price`.
- It creates a deferred nested linked entity for the category of the product. The properties for the category aren't requested until they're needed.

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

The following code sample shows the helper function to create a category linked entity cell value. This function is called by the `contosoLoadService` load service function to create a linked entity for a specific category ID.

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

The following code sample shows the helper function to create a supplier linked entity cell value. This function is called by the `contosoLoadService` load service function to create a linked entity for a specific supplier ID.

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

### Implement the load service function using the helper functions

The following code shows a linked entity load service function that calls the helper functions to create the linked entity cell values. Note the following about the code.

- The load service function parses the incoming `LinkedEntityLoadServiceRequest` object to extract the domain ID and entity IDs. These IDs are used to identify which data domain an entity belongs to, so that the appropriate helper functions can be called to create a linked entity cell value.
- Helper functions create the complete `LinkedEntityCellValue` objects with all properties populated.
- The load service function returns a `LinkedEntityLoadServiceResult` object containing the linked entity cell values in the same order they were requested.

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

# [Load API](#tab/load-api)

### Register and create an event handler

To identify when a linked entity cell value is loaded, you must register a function to handle the `onLinkedEntityCellValueLoaded` event. When the `loadLinkedEntityCellValue` method completes loading the linked entity cell value, the `onLinkedEntityCellValueLoaded` event occurs and a [LinkedEntityCellValueLoadedEventArgs](/javascript/api/excel/excel.linkedentitycellvalueloadedeventargs) object is passed to the event handler. The `LinkedEntityCellValueLoadedEventArgs` object contains the loaded `LinkedEntityCellValue` that you can pass back to the load service function.

The following code shows how to register and implement the event handler for the `onLinkedEntityCellValueLoaded` event.

```typescript
// Map to track pending entity load requests and their resolvers.
const pendingEntityLoads: Map<string, (value: Excel.LinkedEntityCellValue) => void> = new Map();

/**
 * Registers an event handler for the onLinkedEntityCellValueLoaded event.
 * This event occurs when a linked entity cell value has been loaded.
 */
async function registerEvent() {
    await Excel.run(async (context) => {
        const linkedEntityDataDomains = context.workbook.linkedEntityDataDomains;

        // Register the event handler.
        linkedEntityDataDomains.onLinkedEntityCellValueLoaded.add(handleLinkedEntityLoaded);

        await context.sync();
        console.log("Event handler registered successfully. You'll be notified when linked entities are loaded.");
    });
}

/**
 * Event handler that's called when a linked entity cell value is loaded.
 * Extracts the loaded LinkedEntityCellValue and resolves the pending request.
 * @param event - The event object containing the loaded LinkedEntityCellValue.
 */
function handleLinkedEntityLoaded(event: Excel.LinkedEntityCellValueLoadedEventArgs) {
    if (event.linkedEntityCellValue) {
        // Create a unique key for this entity based on entity ID, domain ID, service ID, and culture.
        const entityKey = `${event.linkedEntityCellValue.id.entityId}_${event.linkedEntityCellValue.id.domainId}_${event.linkedEntityCellValue.id.serviceId}_${event.linkedEntityCellValue.id.culture}`;
        
        // Retrieve the resolver function that was stored in pendingEntityLoads.
        // The resolver is a callback from the load service function that waits for the entity to load.
        const resolver = pendingEntityLoads.get(entityKey);
        if (resolver) {
            // Call the resolver with the loaded entity data.
            // This fulfills the Promise that the load service function was waiting on, allowing it to continue and add this entity to the results array.
            resolver(event.linkedEntityCellValue);
            // Clean up the map entry since we no longer need this resolver.
            pendingEntityLoads.delete(entityKey);
        }
    }
}
```

### Implement the load service function using the `loadLinkedEntityCellValue` API

Once you register and implement an event handler to identify when a linked entity cell value is loaded, call the `loadLinkedEntityCellValue` method for each entity in the load service function. Note the following about the code.

- The function must be tagged with `@linkedEntityLoadService`, so that Excel knows where to send the [LinkedEntityLoadServiceRequest](/javascript/api/excel/excel.linkedentityloadservicerequest) object.
- Because the `loadLinkedEntityCellValue` method requires a `LinkedEntityId` as a parameter, you must first create a `LinkedEntityId` object for each entity. To create a `LinkedEntityId` object, use the entity's ID and the domain ID from the `LinkedEntityLoadServiceRequest` object.
- The `loadLinkedEntityCellValue` method is asynchronous and returns results via the `onLinkedEntityCellValueLoaded` event. To preserve the entity order in the `LinkedEntityLoadServiceResult` object so it matches the order in the `LinkedEntityLoadServiceRequest` object, the code uses promises and `Promise.all` to wait for each loaded value.
- Once the linked entity cell value finishes loading, the `onLinkedEntityCellValueLoaded` event occurs and the `handleLinkedEntityLoaded` handler runs to get the linked entity cell value.
- The load service function returns a [LinkedEntityLoadServiceResult](/javascript/api/excel/excel.linkedentityloadserviceresult) object back to Excel to resolve or refresh the linked entity cell values. Note that the load service function must return the loaded linked entity cell values in the same order as their entities in the `LinkedEntityLoadServiceRequest` object.

```typescript
// Linked entity data domain constants
const productsDomainId = "products";
const categoriesDomainId = "categories";
const suppliersDomainId = "suppliers";

// Linked entity cell value constants
const addinDomainServiceId = 268436224;

// Reuse pendingEntityLoads declared in the earlier event-handler snippet.

/**
 * Custom function which acts as the "service" or the data provider for a `LinkedEntityDataDomain`, that is
 * called on demand by Excel to resolve/refresh `LinkedEntityCellValue`s of that `LinkedEntityDataDomain`.
 * @customfunction
 * @linkedEntityLoadService
 * @param {any} request Request to resolve/refresh `LinkedEntityCellValue` objects.
 * @return {any} Resolved/Refreshed `LinkedEntityCellValue` objects that were requested in the passed-in request.
 */
async function contosoLoadService(request: any): Promise<any> {
    const notAvailableError = new CustomFunctions.Error(CustomFunctions.ErrorCode.notAvailable);
    console.log(`Fetching linked entities from request: ${request} ...`);

    try {
        // Parse the request that was passed-in by Excel.
        const parsedRequest: Excel.LinkedEntityLoadServiceRequest = JSON.parse(request);
        // Initialize result to populate and return to Excel.
        const result: Excel.LinkedEntityLoadServiceResult = { entities: [] };

        // Create promises for loading each entity in the request.
        // Promises are needed because the loadLinkedEntityCellValue method is asynchronous.
        // Instead of returning the values immediately, it triggers the onLinkedEntityCellValueLoaded event once the linked entity cell value completes loading.
        const loadPromises: Promise<Excel.LinkedEntityCellValue>[] = [];

        for (const linkedEntityRequest of parsedRequest.entities) {
            // Create a LinkedEntityId object from the request data.
            const linkedEntityId: Excel.LinkedEntityId = {
                entityId: linkedEntityRequest.entityId,
                domainId: parsedRequest.domainId,
                serviceId: addinDomainServiceId,
                culture: linkedEntityRequest.culture
            };

            // Create a promise for this entity that will be resolved when the event handler receives the loaded entity.
            // The resolver (a callback function) is stored in pendingEntityLoads so the event handler can call it later.
            // For production code, consider handling duplicate in-flight requests for the same linked entity key.
            const loadPromise = new Promise<Excel.LinkedEntityCellValue>((resolve, reject) => {
                const entityKey = `${linkedEntityId.entityId}_${linkedEntityId.domainId}_${linkedEntityId.serviceId}_${linkedEntityId.culture}`;
                // Store the resolver callback in the map using a unique key based on entity ID, domain ID, service ID, and culture.
                // The event handler will retrieve this resolver when the entity finishes loading.
                pendingEntityLoads.set(entityKey, resolve);

                // Call loadLinkedEntityCellValue to load the linked entity cell value.
                // This call returns immediately; the actual data arrives via the onLinkedEntityCellValueLoaded event.
                Excel.run(async (context) => {
                    const linkedEntityDataDomains = context.workbook.linkedEntityDataDomains;
                    linkedEntityDataDomains.loadLinkedEntityCellValue(linkedEntityId);
                    await context.sync();
                }).catch((error) => {
                    pendingEntityLoads.delete(entityKey);
                    reject(error);
                });
            });

            loadPromises.push(loadPromise);
        }

        // Wait for all entities to be loaded in order using Promise.all.
        // Promise.all returns results in the same order as the input promises,
        // ensuring that the entities are returned in the same order they arrived in the LinkedEntityLoadServiceRequest object.
        // This also ensures we don't return results until all entities have been resolved by their event handlers.
        const loadedEntities = await Promise.all(loadPromises);

        // Add loaded entities to the result in the same order they were requested.
        // The order is preserved because Promise.all returns results in the same order as the input promises.
        for (const loadedEntity of loadedEntities) {
            result.entities.push(loadedEntity);
        }

        return result;
    } catch (error) {
        console.error(error);
        throw notAvailableError;
    }
}
```

---

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

When you register a data domain, users can refresh it manually at any time, such as by choosing **Data** > **Refresh All**. You can also configure one or more of the following refresh modes.

- `manual`: Refreshes data only when the user chooses to refresh. This is the default mode. Manual refresh is always available, even when the refresh mode is set to `onLoad` or `periodic`.
- `onLoad`: Refreshes data when the data domain is registered, which typically happens when the add-in loads. After that, users refresh the data manually. If you want to refresh data when the workbook opens, configure your add-in to load on document open. For more information, see [Run code in your Office Add-in when the document opens](../develop/run-code-on-document-open.md).
- `periodic`: Refreshes data when the data domain is registered and then refreshes it again at a specified interval. For example, you can refresh every 300 seconds, which is the minimum value. Excel rounds the interval up to the nearest minute because it refreshes only in whole-minute increments.

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
    console.log(`Linked entity domain refreshed: ${eventArgs.id}`);
    console.log(`Refresh status: ${eventArgs.refreshed}`);
    console.log(`Refresh error: ${eventArgs.errors}`);
    return null;
}
```

## Error handling with the linked entity load service

When Excel calls your add-in to get data for a linked entity cell value, an error can occur. If Excel can't connect to your add-in at all, such as when the add-in isn't loaded, Excel displays the `#CONNECT!` error to the user.

If your linked entity load service function encounters an error, it should throw a `CustomFunctions.Error` with `CustomFunctions.ErrorCode.notAvailable`. This causes Excel to show `#CONNECT!` to the user.

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

You can debug most linked entity add-in functionality by following the guidance in [Overview of debugging Office Add-ins](../testing/debug-add-ins-overview.md). However, the linked entity load service function can run in a shared runtime or in a JavaScript-only runtime, also known as a custom functions runtime. If you implement the function in a JavaScript-only runtime, use [Custom functions debugging in a non-shared runtime](custom-functions-debugging.md).

The linked entity load service function uses the custom functions architecture, regardless of which runtime you use. However, there are significant differences from regular custom functions.

Linked entity load service functions have the following differences from custom functions:

- They don't appear to end users for usage in formulas.
- They don't support the JSDoc tags `@streaming` or `@volatile`. The user will see a **#CALC!** error if these tags are used.

Linked entity load service functions have the following similarities with custom functions:

- They use [Custom functions naming and localization](custom-functions-naming.md).
- They use the same error handling approach.

## Behavior in Excel 2019 and earlier

If someone opens a worksheet with linked entity cell values on an older version of Excel that doesn't support linked entity cell values, Excel shows the cell values as errors. This behavior is by design. This behavior is also why you set the `basicType` to `Error` and the `basicValue` to `#VALUE!` every time you insert or update a linked entity cell value. This error is the fallback that Excel uses on older versions.

## Best practices

- Don't use exclamation marks in the `entityID` or `domainId` values.
- Register linked entity data domains in your `Office.onReady` initialization code so users can refresh linked entity cell values as soon as the add-in loads.
- After publishing your add-in, don't change the linked entity data domain IDs. Consistent IDs across the same logical objects helps with performance.
- Always provide the `text` property when creating a new linked entity cell value. This value is displayed while Excel calls your data provider function to get the remaining property values. Otherwise, the user sees a blank cell until the data is retrieved.

## See also

- [Overview of data types in Excel add-ins](excel-data-types-overview.md)
- [Use data types in Excel add-ins](excel-data-types-concepts.md)
- [Use cards with cell value data types](excel-data-types-entity-card.md)
- [Add properties to Excel basic cell values](excel-data-types-add-properties-to-basic-cell-values.md)
- [Custom functions and data types](custom-functions-data-types-concepts.md)
