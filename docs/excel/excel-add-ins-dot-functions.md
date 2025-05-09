---
title: Add reference methods to cell values
description: Learn how to add reference methods to cell values.
ms.date: 04/29/2025
ms.localizationpriority: medium
---

# Add reference methods to cell values

Add reference methods to cell values to provide the user access to dynamic calculations based on the cell value. The `EntityCellValue` and `LinkedEntityCellValue` types support reference methods. For example, add a method to a product entity value that converts its weight to different units.

The following screenshot shows an example of adding a `ConvertWeight` method to a product entity value representing pancake mix.

:::image type="content" source="../images/excel-add-in-dot-function.png" alt-text="Screenshot of Excel formula showing =A1.ConvertWeight( ounces )":::

The `DoubleCellValue`, `BooleanCellValue`, and `StringCellValue` types also support reference methods. The following screenshot shows an example of adding a `ConvertToRomanNumeral` method to a double value type.

:::image type="content" source="../images/excel-add-in-dot-function-roman-numeral.png" alt-text="Screenshot of Excel formula showing =A1.ConvertToRomanNumeral()":::

Reference methods don’t appear on the data type card for the user.

:::image type="content" source="../images/excel-add-in-dot-function-data-card.png" alt-text="Screenshot of data card for Pancake mix data type, but no reference methods are listed.":::

## Add a reference method to an entity value

To add a reference method to an entity value, describe it in the JSON using the `Excel.JavaScriptCustomFunctionReferenceCellValue` type. The following code sample shows how to define a simple method that returns the value 27.

```typescript
const referenceCustomFunctionGet27: Excel.JavaScriptCustomFunctionReferenceCellValue = { 
  type: Excel.CellValueType.function,
  functionType: Excel.FunctionCellValueType.javaScriptReference,
  namespace: "CONTOSO", 
  id: "GET27" 
} 
```

The properties are described in the following table.

|Property  |Description  |
|---------|---------|
|**type** | Specifies the type of reference. This property only supports `function` and must be set to `Excel.CellValueType.function`.|
|**functionType**     | Specifies the type of function. This property only supports JavaScript reference functions, and must be set to `Excel.FunctionCellValueType.javaScriptReference`.|
|**namespace** | The namespace that contains the custom function. This value must match the namespace specified by the [customFunctions.namespace element](/microsoftteams/platform/resources/schema/manifest-schema-dev-preview) in the unified manifest, or the [Namespace element](/javascript/api/manifest/namespace) in the add-in only manifest.|
|**id** | The name of the custom function to map to this reference method. The name is the uppercase version of the custom function name. |

When you create the entity value, add the reference method to the properties list. The following code sample shows how to create a simple entity value named `Math`, and add a reference method to it. `Get27` is the method name that will appear to the user. For example `A1.Get27()`.

```typescript
function makeMathEntity(value: number){
  const entity: Excel.EntityCellValue = {
    type: Excel.CellValueType.entity,
    text: "Math value",
    properties: {
      "value": {
        type: Excel.CellValueType.double,
        basicValue: value,
        numberFormat: "#"
      },
      Get27: referenceCustomFunctionGet27
    }
  };
  return entity;
}
```

The following code sample shows how to create an instance of the `Math` entity and add it to the selected cell.

```typescript
// Add entity to selected cell.
async function addEntityToCell(){
  const entity: Excel.EntityCellValue = makeMathEntity(10);
  await Excel.run( async (context) => {
    const cell = context.workbook.getActiveCell();
    cell.valuesAsJson = [[entity]];
    await context.sync();
  });
}
```

Finally, the reference method is implemented by a custom function. The following code sample shows how to implement the custom function.

```typescript
/**
 * Returns the value 27.
 * @customfunction
 * @excludeFromAutocomplete
 * @returns {number} 27
 */
function get27() {
  return 27;
}
```

In the previous code sample, the `@excludeFromAutocomplete` tag ensures the custom function doesn't appear to the user in the Excel UI when entering it in a search box. However, note that a user can still call the custom function separately from an entity value if they enter it directly into a cell.

When the code runs, it creates a `Math` entity value as shown in the following screenshot. The method appears in formula autocomplete when the user references the entity value from a formula.

:::image type="content" source="../images/excel-data-types-reference-method-autocomplete.png" alt-text="Screenshot of entering A1. in Excel and formula autocomplete displaying the Get27 reference method.":::

## Add arguments

If your reference method needs arguments, add them to the custom function. The following code example shows how to add an argument named `x` to a method named `AddValue`. The method adds one to the `x` value by calling a custom function named `AddValue`.

```typescript
/**
 * Adds a value to 1.
 * @customfunction
 * @excludeFromAutocomplete
 * @param {number} x The value to add to 1.
 * @return {number[][]}  Sum of x and 1.
 */
function addValue(x): number[][] {  
  return [[x+1]];
}
```

## Reference the entity value as a calling object

A common scenario is that your methods need to reference properties on the entity value itself to perform calculations. For example, it's more useful if the `AddValue` method adds the argument value to the entity value itself. Specify that the entity value be passed in as the first argument by applying the `@capturesCallingObject` tag to the custom function as shown in the following code example.

```typescript
/**
 * Adds x to the calling object.
 * @customfunction
 * @excludeFromAutocomplete
 * @capturesCallingObject
 * @param {any} math The math object (calling object).
 * @param {number} x The value to add.
 * @return {number[][]}  Sum.
 */
function addValue(math, x): number[][] {  
  const result: number = math.properties["value"].basicValue + x;
  return [[result]];
}
```

Note that the argument name can be whatever you decide, as long as it conforms to the Excel syntax rules as specified in [Names in formulas](https://support.microsoft.com/office/names-in-formulas-fc2935f9-115d-4bef-a370-3aa8bb4c91f1). Since we know this is a math entity, we name the calling object argument `math`. The argument name can be used in the body calculation. In the previous code sample, it retrieves the `math.[value]` property to perform the calculation.

The following code sample shows the implementation of the `Contoso.AddValue` function.

```typescript
/**
 * Adds x to the calling object.
 * @customfunction
 * @excludeFromAutocomplete
 * @param {any} math The math object (calling object).
 * @param {number} x The value to add.
 * @return {number[][]}  Sum.
 */
function addValue(math, x): number[][] {  
  const result: number = math.properties["value"].basicValue + x;
  return [[result]];
}
```

Note the following about the previous code sample.

- The `@excludeFromAutocomplete` tag ensures the custom method doesn't appear to the user in the Excel UI when entering it in a search box. However, note that a user can still call the custom function separately from an entity value if they enter it directly into a cell.
- The calling object is always passed as the first argument and must by of type `any`. In this case, it's named `math` and is used to get the value property from the `math` object.
- It returns a double array of numbers.
- When the user interacts with the reference method in Excel, they don’t see the calling object as an argument.

## Example: Calculate product sales tax

The following code shows how to implement a custom function that calculates the sales tax for the unit price of a product.

```typescript
/**
 * Calculates the price when a sales tax rate is applied.
 * @customfunction
 * @excludeFromAutocomplete
 * @capturesCallingObject
 * @param {any} product The product entity value (calling object).
 * @param {number} taxRate The tax rate (0.11 = 11%).
 * @return {number[][]}  Product unit price with tax rate applied.
 */
function applySalesTax(product, taxRate): number[][] {
  const unitPrice: number = product.properties["Unit Price"].basicValue;
  const result: number = unitPrice * taxRate + unitPrice;
  return [[result]];
}
```

The following code sample shows how to specify the reference method and includes the `id` of the `applySalesTax` custom function.

```typescript
const referenceCustomFunctionCalculateSalesTax: Excel.JavaScriptCustomFunctionReferenceCellValue = { 
  type: Excel.CellValueType.function,
  functionType: Excel.FunctionCellValueType.javaScriptReference,
  namespace: "CONTOSO", 
  id: "APPLYSALESTAX" 
} 
```

The following code shows how to add the reference method to the `product` entity value.

```typescript
function makeProductEntity(productID: number, productName: string, price: number) {
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
      "Unit Price": {
        type: Excel.CellValueType.formattedNumber,
        basicValue: price,
        numberFormat: "$* #,##0.00"
      },
      applySalesTax: referenceCustomFunctionCalculateSalesTax
    },
  };
  return entity;
}
```

## Exclude custom functions from the Excel UI

Use the `@excludeFromAutoComplete` tag in the comments description of custom functions used by reference methods to indicate that the function will be excluded from the autocomplete drop-down list and Formula Builder. This helps prevent the user from accidentally using a custom function separately from its entity value.

> [!NOTE]
> If the function is manually entered correctly in the grid, the function will still execute. Also, a function can’t have both `@excludeFromAutoComplete` and `@linkedEntityLoadService` tags.

The `@excludeFromAutoComplete` tag is processed during build to generate a **functions.json** file by the **Custom-Functions-Metadata** package. This package is automatically added to the build process if you start with yo office and choose a custom function template. If you aren't using this package, you'll need to add the `excludeFromAutoComplete` property manually to the **functions.json** file.

The following code sample shows how to manually describe the `APPLYSALESTAX` with JSON in the **functions.json** file. The `excludeFromAutoComplete` property is set to `true`.

```typescript
{
    "description": "Calculates the price when a sales tax rate is applied.",
    "id": "APPLYSALESTAX",
    "name": "APPLYSALESTAX",
    "options": {
        "excludeFromAutoComplete": true,
        "capturesCallingObject": true
    },
    "parameters": [
        {
            "description": "The product entity value (calling object).",
            "name": "product",
            "type": "any"
        },
        {
            "description": "The tax rate (0.11 = 11%).",
            "name": "taxRate",
            "type": "number"
        }
    ],
    "result": {
        "dimensionality": "matrix",
        "type": "number"
    }
},
```

For more information, see [Manually create JSON metadata for custom functions](custom-functions-json.md).

## Add a function to a basic value type

To add functions to the basic value types of `Boolean`, `double`, and `string`, the process is the same as for entity values. Describe the function with JSON as a reference method. The following code sample shows how to create a double basic value with a function `AddValue()` that adds a value `x` to the basic value.

```typescript
/**
 * Adds the value x to the number value.
 * @customfunction
 * @capturesCallingObject
 * @param {any} numberValue The number value (calling object).
 * @param {number} x The value to add to 1.
 * @return {number[][]}  Sum of x and 1.
 */
export function addValue(numberValue: any, x: number): number[][] {
  return [[x+numberValue.basicValue]];
}

```

The following code sample shows how to add the `addValue` reference method to a simple number in Excel.

```typescript
const referenceCustomFunctionAddValue: Excel.JavaScriptCustomFunctionReferenceCellValue = { 
  type: Excel.CellValueType.function,
  functionType: Excel.FunctionCellValueType.javaScriptReference,
  namespace: "CONTOSO", 
  id: "ADDVALUE" 
} 

async function createSimpleNumber() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange("A1");
    range.valuesAsJson = [
      [
        {
          type: Excel.CellValueType.double,
          basicType: Excel.RangeValueType.double,
          basicValue: 6.0,
          properties: {
            addValue: referenceCustomFunctionAddValue
          }
        }
      ]
    ];
    await context.sync();
  });
}
```

## Optional arguments

The following code sample shows how to create a reference method that accepts optional arguments. The reference method is named `generateRandomRange` and it generates a range of random values.

```typescript
const referenceCustomFunctionOptional: Excel.JavaScriptCustomFunctionReferenceCellValue = { 
  type: Excel.CellValueType.function,
  functionType: Excel.FunctionCellValueType.javaScriptReference,
  namespace: "CONTOSO", 
  id: "GENERATERANDOMRANGE" 
}

function makeProductEntity(productID: number, productName: string, price: number) {
  const entity: Excel.EntityCellValue = {
    type: Excel.CellValueType.entity,
    text: productName,
    properties: {
      "Product ID": {...},
      "Product Name": {...},
      "Unit Price": {...},
      generateRandomRange: referenceCustomFunctionOptional
    },
  };
  return entity;
}
```

The following code sample shows the implementation of the reference method as a custom function named `generateRandomRange`. It returns a dynamic array of random values matching the number of `rows` and `columns` specified. The `min` and `max` arguments are optional, and if not specified will default to `1` and `10`.

```typescript
/**
 * Generate a dynamic array of random numbers.
 * @customfunction
 * @excludeFromAutocomplete
 * @param {number} rows Number of rows to generate.
 * @param {number} columns Number of columns to generate.
 * @param {number} [min] Lowest number that can be generated. Default is 1.
 * @param {number} [max] Highest number that can be generated. Default is 10.
 * @returns {number[][]} A dynamic array of random numbers.
 */
function generateRandomRange(rows, columns, min, max) {
  // Set defaults for any missing optional arguments.
  if (min===null) min = 1;
  if (max === null) max = 10;

  let numbers = new Array(rows);
  for (let r = 0; r < rows; r++) {
    numbers[r] = new Array(columns);
    for (let c = 0; c < columns; c++) {
      numbers[r][c] = Math.round(Math.random() * (max - min) ) + min;
    }
  }
  return numbers;
}
```

When the user enters the function in Excel, autocomplete shows the properties of the function, and indicates optional arguments by surrounding them in brackets []. The following image shows an example of entering optional parameters using the `generateRandomRange` reference method.

:::image type="content" source="../images/excel-data-types-reference-method-optional-parameters.png" alt-text="Screenshot of entering generateRandomRange method in Excel.":::

## Multiple parameters

Reference methods support multiple parameters similar to how the Excel `SUM` function supports multiple parameters. The following code sample shows how to create a reference function that concatenates zero or more product names passed in a products array. The function is shown to the user as `concatProductNames([products], ...)`.

```typescript
/** 
 * @customfunction 
 * @excludeFromAutocomplete 
 * @description Concatenate the names of given products, joined by " | " 
 * @param {any[]} products - The products to concatenate.
 * @returns A string of concatenated product names. 
 */ 
function concatProductNames(products: any[]): string { 
  return products.map((product) => product.properties["Product Name"].basicValue).join(" | "); 
}
```

The following code sample shows how to create an entity with the `concatProductNames` reference method.

```typescript
const referenceCustomFunctionMultiple: Excel.JavaScriptCustomFunctionReferenceCellValue = { 
  type: Excel.CellValueType.function,
  functionType: Excel.FunctionCellValueType.javaScriptReference,
  namespace: "CONTOSO", 
  id: "CONCATPRODUCTNAMES" 
} 

function makeProductEntity(productID: number, productName: string, price: number) {
  const entity: Excel.EntityCellValue = {
    type: Excel.CellValueType.entity,
    text: productName,
    properties: {
      "Product ID": {...},
      "Product Name": {...},
      "Unit Price": {...},
      concatProductNames: referenceCustomFunctionMultiple,
    },
  };
  return entity;
}
```

The following image shows an example of entering multiple parameters using the `concatProductNames` reference method.

:::image type="content" source="../images/excel-data-types-reference-method-repeating-parameters.png" alt-text="Screenshot of entering concatProductNames method in Excel passing A1 and A2 which contain a bicycle and unicycle product entity value.":::

### Multiple parameters with ranges

To support passing ranges to your reference method such as **B1:B3**, use a multidimensional array. The following code sample shows how to create a reference function that sums zero or more parameters which can include ranges.

```typescript
/** 
 * @customfunction 
 * @excludeFromAutocomplete 
 * @description Calculate the sum of arbitrary parameters. 
 * @param {number[][][]} operands - The operands to sum. 
 * @returns  The sum of all operands. 
 */ 
function sumAll(operands: number[][][]): number { 
  let total: number = 0; 
 
  operands.forEach(range => { 
    range.forEach(row => { 
      row.forEach(num => { 
        total += num; 
      }); 
    }); 
  }); 
 
  return total; 
} 
```

The following code sample shows how to create an entity with the `sumAll` reference method.

```typescript
const referenceCustomFunctionRange: Excel.JavaScriptCustomFunctionReferenceCellValue = { 
  type: Excel.CellValueType.function,
  functionType: Excel.FunctionCellValueType.javaScriptReference,
  namespace: "CONTOSO", 
  id: "SUMALL" 
} 

function makeProductEntity(productID: number, productName: string, price: number) {
  const entity: Excel.EntityCellValue = {
    type: Excel.CellValueType.entity,
    text: productName,
    properties: {
      "Product ID": {...},
      "Product Name": {...},
      "Unit Price": {...},
      sumAll: referenceCustomFunctionRange
    },
  };
  return entity;
}
```

The following image shows an example of entering multiple parameters including a range parameter using the `sumAll` reference method.

:::image type="content" source="../images/excel-data-types-reference-method-optional-ranges.png" alt-text="Screenshot of entering sumAll method in Excel passing an optional range of B1:B2.":::

## Support details

Reference methods are supported in all custom function types, such as [volatile](custom-functions-volatile.md), and [streaming](custom-functions-web-reqs.md) custom functions. Also, all custom function return types are supported ([matrix, scalar, and error](custom-functions-json-autogeneration.md))
A linked entity can’t have a custom function that combines both a reference method and a data provider. Be sure when developing linked entities to keep these types of custom functions separate.

## See also

- [Create custom functions in Excel](custom-functions-overview.md)
- [Excel.CellValue](/javascript/api/excel/excel.cellvalue)
- [LinkedEntityCellValue](/javascript/api/excel/excel.linkedentitycellvalue)
- [DoubleCellValue](/javascript/api/excel/excel.doublecellvalue)
- [StringCellValue](/javascript/api/excel/excel.stringcellvalue)
- [BooleanCellValue](/javascript/api/excel/excel.booleancellvalue)
