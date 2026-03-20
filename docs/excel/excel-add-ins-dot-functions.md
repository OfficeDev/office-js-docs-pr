---
title: Add reference methods to cell values
description: Learn how to add reference methods to cell values.
ms.date: 11/04/2025
ms.localizationpriority: medium
---

# Add reference methods to cell values

Add reference methods to cell values to give users access to dynamic calculations based on the cell value. The [`EntityCellValue`](/javascript/api/excel/excel.entitycellvalue) and [`LinkedEntityCellValue`](/javascript/api/excel/excel.linkedentitycellvalue) types support reference methods. For example, add a method to a product entity value that converts its weight to different units.

The following image shows an example of adding a `ConvertWeight` method to a product entity value representing pancake mix.

:::image type="content" source="../images/excel-add-in-dot-function.png" alt-text="Excel formula showing =A1.ConvertWeight(ounces).":::

The [`DoubleCellValue`](/javascript/api/excel/excel.doublecellvalue), [`BooleanCellValue`](/javascript/api/excel/excel.booleancellvalue), and [`StringCellValue`](/javascript/api/excel/excel.stringcellvalue) types also support reference methods. The following image shows an example of adding a `ConvertToRomanNumeral` method to a double value type.

:::image type="content" source="../images/excel-add-in-dot-function-roman-numeral.png" alt-text="Excel formula showing =A1.ConvertToRomanNumeral()":::

Reference methods don't appear on the data type card.

:::image type="content" source="../images/excel-add-in-dot-function-data-card.png" alt-text="Data card for the Pancake mix data type, but no reference methods are listed.":::

## Add a reference method to an entity value

To add a reference method to an entity value, define it in JSON by using the `Excel.JavaScriptCustomFunctionReferenceCellValue` type. The following code sample shows how to define a simple method that returns the value 27.

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

When you create the entity value, add the reference method to the properties list. The following code sample shows how to create a simple entity value named `Math` and add a reference method to it. `Get27` is the method name that appears to users (for example: `A1.Get27()`).

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

Finally, implement the reference method with a custom function. The following code sample shows how to implement the custom function.

```typescript
/**
 * Returns the value 27.
 * @customfunction
 * @excludeFromAutoComplete
 * @returns {number} 27
 */
function get27() {
  return 27;
}
```

In the previous code sample, the `@excludeFromAutoComplete` tag ensures the custom function doesn't appear in the Excel UI when a user enters it in a search box. However, a user can still call the custom function separately from an entity value if they enter it directly into a cell.

When the code runs, it creates a `Math` entity value as shown in the following image. The method appears in formula AutoComplete when the user references the entity value from a formula.

:::image type="content" source="../images/excel-data-types-reference-method-autocomplete.png" alt-text="Entering 'A1.' in Excel with formula AutoComplete displaying the 'Get27' reference method.":::

## Add arguments

If your reference method needs arguments, add them to the custom function. The following code example shows how to add an argument named `x` to a method named `addValue`. The method adds one to the `x` value by calling a custom function named `addValue`.

```typescript
/**
 * Adds a value to 1.
 * @customfunction
 * @excludeFromAutoComplete
 * @param {number} x The value to add to 1.
 * @return {number[][]}  Sum of x and 1.
 */
function addValue(x): number[][] {  
  return [[x+1]];
}
```

## Reference the entity value as a calling object

A common scenario is that your methods need to reference properties on the entity value itself to perform calculations. For example, it's more useful if the `addValue` method adds the argument value to the entity value itself. Specify that the entity value is passed as the first argument by applying the `@capturesCallingObject` tag to the custom function as shown in the following code example.

```typescript
/**
 * Adds x to the calling object.
 * @customfunction
 * @excludeFromAutoComplete
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

You can use any argument name that conforms to the Excel syntax rules in [Names in formulas](https://support.microsoft.com/office/names-in-formulas-fc2935f9-115d-4bef-a370-3aa8bb4c91f1). Because this is a math entity, the calling object argument is named `math`. The argument name can be used in the calculation.

Note the following about the previous code sample.

- The `@excludeFromAutoComplete` tag ensures that the custom function doesn't appear in the Excel UI when a user enters it in a search box. However, a user can still call the custom function separately from an entity value if they enter it directly into a cell.
- The calling object is always passed as the first argument and must be of type `any`. In this case, it's named `math` and is used to get the value property from the `math` object.
- It returns a double array of numbers.
- When the user interacts with the reference method in Excel, they don't see the calling object as an argument.

## Example: Calculate product sales tax

The following code shows how to implement a custom function that calculates the sales tax for the unit price of a product.

```typescript
/**
 * Calculates the price when a sales tax rate is applied.
 * @customfunction
 * @excludeFromAutoComplete
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

Use the `@excludeFromAutoComplete` tag in the JSDoc tag of custom functions used by reference methods to indicate that the function be excluded from formula AutoComplete and Formula Builder. This helps prevent users from accidentally using a custom function separately from its entity value.

> [!NOTE]
> If the function is manually entered correctly in the grid, the function still runs.

>[!IMPORTANT]
> A function can't have both `@excludeFromAutoComplete` and `@linkedEntityLoadService` tags.

The `@excludeFromAutoComplete` tag is processed during build to generate a **functions.json** file by the **Custom-Functions-Metadata** package. This package is automatically added to the build process if you create your add-in with the Yeoman generator for Office Add-ins and choose a custom functions template. If you aren't using the **Custom-Functions-Metadata** package, you'll need to add the `excludeFromAutoComplete` property manually to the **functions.json** file.

The following code sample shows how to manually define the `APPLYSALESTAX` custom function with JSON in the **functions.json** file. The `excludeFromAutoComplete` property is set to `true`.

```json
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

To add functions to the basic value types of `Boolean`, `double`, and `string`, use the same process as you would for entity values. The following code sample shows how to create a double basic value with a custom function called `addValue`. The function adds the value `x` to the basic value.

```typescript
/**
 * Adds the value x to the number value.
 * @customfunction
 * @capturesCallingObject
 * @param {any} numberValue The number value (calling object).
 * @param {number} x The value to add.
 * @return {number[][]}  Sum of the number value and x.
 */
export function addValue(numberValue: any, x: number): number[][] {
  return [[x+numberValue.basicValue]];
}

```

The following code sample shows how to define the `addValue` custom function from the preceding sample in JSON and then reference it with a method called `createSimpleNumber`.

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
 * Generates a dynamic array of random numbers.
 * @customfunction
 * @excludeFromAutoComplete
 * @param {number} rows Number of rows to generate.
 * @param {number} columns Number of columns to generate.
 * @param {number} [min] Lowest number that can be generated. Default is 1.
 * @param {number} [max] Highest number that can be generated. Default is 10.
 * @returns {number[][]} A dynamic array of random numbers.
 */
function generateRandomRange(rows, columns, min, max) {
  // Set defaults for any missing optional arguments.
  if (min === undefined) min = 1;
  if (max === undefined) max = 10;

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

When the user enters the custom function in Excel, AutoComplete shows the properties of the function and indicates optional arguments by surrounding them in brackets (`[]`). The following image shows an example of entering optional parameters by using the `generateRandomRange` reference method.

:::image type="content" source="../images/excel-data-types-reference-method-optional-parameters.png" alt-text="Screenshot of entering generateRandomRange method in Excel.":::

## Multiple parameters

Reference methods support multiple parameters, similar to how the Excel `SUM` function supports multiple parameters. The following code sample shows how to create a reference function that concatenates zero or more product names passed in a products array. The function is shown to the user as `concatProductNames([products], ...)`.

```typescript
/** 
 * @customfunction 
 * @excludeFromAutoComplete 
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

The following image shows an example of entering multiple parameters by using the `concatProductNames` reference method.

:::image type="content" source="../images/excel-data-types-reference-method-repeating-parameters.png" alt-text="Screenshot of entering concatProductNames method in Excel passing A1 and A2 which contain a bicycle and unicycle product entity value.":::

### Multiple parameters with ranges

To support passing ranges to your reference method such as **B1:B3**, use a multidimensional array. The following code sample shows how to create a reference function that sums zero or more parameters which can include ranges.

```typescript
/** 
 * @customfunction 
 * @excludeFromAutoComplete 
 * @description Calculate the sum of arbitrary parameters. 
 * @param {number[][][]} operands - The operands to sum. 
 * @returns The sum of all operands. 
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

The following image shows an example of entering multiple parameters, including a range parameter, by using the `sumAll` reference method.

:::image type="content" source="../images/excel-data-types-reference-method-optional-ranges.png" alt-text="Screenshot of entering sumAll method in Excel passing an optional range of B1:B2.":::

## Support details

Reference methods are supported in all custom function types, such as [volatile](custom-functions-volatile.md) and [streaming](custom-functions-web-reqs.md#make-a-streaming-function) functions. All custom function return types&mdash;matrix, scalar, and error&mdash;are supported.

> [!IMPORTANT]
> A linked entity can't have a custom function that combines both a reference method and a data provider. When you develop linked entities, keep these types of custom functions separate.

## See also

- [Create custom functions in Excel](custom-functions-overview.md)
- [Autogenerate JSON metadata for custom functions](custom-functions-json-autogeneration.md)
- [Manually create JSON metadata for custom functions](custom-functions-json.md)
