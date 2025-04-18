---
title: Add lambda methods to cell values
description: Add lambda methods to cell values to support dot syntax in Excel formulas.
ms.date: 04/16/2025
ms.localizationpriority: medium
---

# Add lambda methods to cell values

Add lambda methods to cell values to provide the user access to dynamic calculations based on the cell value. You can add lambda methods to the `EntityCellValue` and `LinkedEntityCellValue` types. For example, add a method to a product entity value that converts its weight to different units.

The following screenshot shows an example of adding a `ConvertWeight` method to a product entity value representing pancake mix.

:::image type="content" source="/images/excel-add-in-dot-function.png" alt-text="Screenshot of Excel formula showing =A1.ConvertWeight( ounces )":::

You can also add lambda methods to the `DoubleCellValue`, `BooleanCellValue`, and `StringCellValue` types. The following screenshot shows an example of adding a `ConvertToRomanNumeral` method to a double value type.

:::image type="content" source="/images/excel-add-in-dot-function-roman-numeral.png" alt-text="Screenshot of Excel formula showing =A1.ConvertToRomanNumeral()":::

Lambda methods don’t appear on the data type card for the user.

:::image type="content" source="/images/excel-add-in-dot-function-data-card.png" alt-text="Screenshot of data card for Pancake mix data type, but no lamdba methods are listed.":::

## Add a lambda method to an entity value

To add a lambda method to an entity value, describe it in JSON using the `LambdaCellValue` type. The following code sample shows how to define a simple method that returns the value 27.

```javascript
const lambdaReturn27 = {
  type: Excel.CellValueType.entity,
  body: "Contoso.get27()",
  callingObjectAsFirstArgument: false,
  shortDescription: "Returns the value 27",
  longDescription: "Returns the value 27",
  example: "=A1.Get27()",
  arguments: [
  ]
}
```

The properties are described as follows.

|Property  |Description  |
|---------|---------|
|**type** | Must be Lambda. The type is of Excel.CellValueType.lambda.|
|**body**     |The name of the custom function to call with arguments.|
|**callingObjectAsFirstArgument** | Whether the method is passed the calling object as the first argument. The calling object is the entity value the method is operating on. If not specified, this property defaults to true.|
|**shortDescription** | Short description of the method displayed to the user. Note: At this time, this property is not used by Excel. Optional. |
|**longDescription** | Long description of the method displayed to the user. Note: At this time, this property is not used by Excel. Optional. |
| **example** | Shows the user an example of how to call the method. Note: At this time, this property is not used by Excel. Optional.  |

When you create the entity value, add the lambda method to the properties list. The following code sample shows how to create a simple entity value named `Math`, and add a lambda method to it. `Get27` is the method name that will appear to the user. For example `A1.Get27()`.

```javascript
function MakeMathEntity(value: number){
  const entity: Excel.EntityCellValue = {
    type: Excel.CellValueType.entity,
    text: "Math value",
    properties: {
      "value": {
        type: Excel.CellValueType.formattedNumber,
        basicValue: value,
        numberFormat: "#"
      },
      Get27: lambdaReturn27
    }
  };
  return entity;
}
```

The following code shows how to create an instance of the `Math` entity and add it to the selected cell.

```javascript
// Add entity to selected cell
async function addEntityToCell(){
  const entity: Excel.EntityCellValue = makeMathEntity(10);
  await Excel.run( async (context) => {
    const cell = context.workbook.getActiveCell();
    cell.valuesAsJson = [[entity]];
    await context.sync();
  });
}
```

Finally, the lambda method needs the custom function to call to implement. Previously the custom function was called as `Contoso.Get27()` in the body property of the lambda method. The following code shows how to implement the custom function.

```javascript

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

In the previous code sample, the `@excludeFromAutocomplete` attribute ensures the custom method doesn't appear to the user in the Excel UI when entering it in a search box. However, note that a user can still call the custom function separately from an entity value if they enter it directly into a cell.

When the code runs it will create a `Math` entity value as shown in the following screenshot. The method appears in formula autocomplete when the user references the entity value from a formula.

## Add arguments

You can also accept arguments to lambda methods. To do this you need to describe the arguments in the JSON for the lambda method (`LambdaCellValue`). The following code shows how to add an argument named `x` to a method named `AddValue`. The method adds one to the `x` value by calling a custom function named `AddValue`.

```javascript
const lambdaOneArgument = {
  type: Excel.CellValueType.lambda,
  body: "Contoso.AddValue(x, 1)",
  callingObjectAsFirstArgument: false,
  shortDescription: "AddValue",
  longDescription: "Adds one to the argument",
  example: "=A1.AddValue(1)",
  arguments: [
    {
      name: "x"
    }
  ]
};
```

## Reference the entity value as a calling object

A common scenario is that your methods need to reference properties on the entity value itself to perform calculations. For example, it's more useful if the AddValue method adds the argument value to the entity value itself. Specify that the entity value be passed in as the first argument by setting `callingObjectAsFirstArgument` to `true`. The following JSON shows how to add the calling object.

```javascript
const lambdaOneArgument = {
  type: Excel.CellValueType.lambda,
  body: "Contoso.AddValue(math, x)",
  callingObjectAsFirstArgument: true,
  shortDescription: "AddValue",
  longDescription: "Adds x to the math value",
  example: "=A1.AddValue(5)",
  arguments: [
    {
      name: "math"
    },
    {
      name: "x"
    }
  ]
};
```

Note that the argument name can be whatever you decide, as long as it conforms to the Excel syntax rules as specified in [Names in formulas](https://support.microsoft.com/office/names-in-formulas-fc2935f9-115d-4bef-a370-3aa8bb4c91f1). Since we know this is a math entity we name the calling object argument `math`. The argument name can be used in the body calculation. In the previous code sample it retrieves the `math.[value]` property to perform the calculation.

The following code sample shows the implementation of the `Contoso.AddValue` function.

```javascript
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

- The `@excludeFromAutocomplete` attribute ensures the custom method doesn't appear to the user in the Excel UI when entering it in a search box. However, note that a user can still call the custom function separately from an entity value if they enter it directly into a cell.
- The calling object is always passed as the first argument and must by of type `any`. In this case, it's named `math` and is used to get the value property from the `math` object.
- It returns a double array of numbers.

When the user interacts with the lambda method in Excel, they don’t see the calling object as an argument.

## Example: Calculate product sales tax

The following code shows how to implement a custom function that calculates the sales tax for the unit price of a product.

```javascript
/**
 * Calculates the price when a sales tax rate is applied.
 * @customfunction
 * @excludeFromAutocomplete
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

Call the custom function from the body property of the lambda method. The following code sample shows how to call the custom function from the `body` property.

```javascript
const lambdaCalculateSalesTax = {
  type: Excel.CellValueType.lambda,
  body: "Contoso.applySalesTax(product, taxRate)",
  callingObjectAsFirstArgument: true, //Set to true to get the calling object
  shortDescription: "Calculates sales tax",
  longDescription: "Calculates the sales tax for the product",
  example: "=A1.ApplySalesTax()",
  arguments: [
    {
      name: "product" // First argument is the calling object.
    },
    {
      name: "taxRate"
    }
  ]
};
```

The following code shows how to add the lambda method to the `product` entity value.

```javascript
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
      applySalesTax: lambdaCalculateSalesTax
    }
  };
  return entity;
}
```

## Exclude custom functions from the Excel UI

We recommend you use the `@excludeFromAutoComplete` tag in the comments description of custom functions used by lambda methods. It indicates that the function will be excluded from the autocomplete drop-down list and Formula Builder. This helps prevent the user from accidentally using a custom function separately from its entity value.

> [!NOTE]
> If the function is manually entered correctly in the grid, the function will still execute. Also, a function can’t have both `@excludeFromAutoComplete` and `@linkedEntityDataProvider` tags.

For the full list of properties in addition to the properties specified by [Manually create JSON metadata for custom functions](custom-functions-json.md) the `excludeFromAutoComplete` property is available for lambda methods.

The `@excludeFromAutoComplete` tag is processed during build to generate a **functions.json** file by the **Custom-Functions-Metadata** package. This package is automatically added to the build process if you start with yo office and choose a custom function template. If you aren't using this package, you'll need to add the `excludeFromAutoComplete` property manually to the **functions.json** file. 

The following example shows how to manually describe the `APPLYSALESTAX` with JSON in the **functions.json** file. The `excludeFromAutoComplete` property is set to `true`.

```javascript
{
            "description": "Calculates the price when a sales tax rate is applied.",
            "id": "APPLYSALESTAX",
            "name": "APPLYSALESTAX",
            "options": {
                "excludeFromAutoComplete": true
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

## Optimize calls

You can improve the performance of lambda methods by only passing the properties of the calling object that you use. To do this, explicitly specify which properties to use when calling your custom function from the `body` of the lambda method. The following code sample shows how to only pass the unit price of the product in the call to `applySalesTax`.

```javascript
const lambdaCalculateSalesTax = {
  type: Excel.CellValueType.lambda,
  body: "Contoso.applySalesTax(product.[Unit Price], taxRate)",
  callingObjectAsFirstArgument: true, //Set to true to get the calling object
  shortDescription: "Apply sales tax",
  longDescription: "Apply sales tax",
  example: "=A1.WithSalesTax(5.5)",
  arguments: [
    {
      name: "product" // First parameter is the calling object.
    },
    {
      name: "taxRate"
    }
  ]
};
```

Then in the custom function, the first argument will still be an `EntityCellValue` type, but it only contains the property you specified. The following code sample shows how to modify the previous custom function to accept the unit price as the first argument (not the entire product entity value). To access the value it uses the `basicValue` property.

```javascript
/**
 * Calculates the price when a sales tax rate is applied.
 * @customfunction
 * @excludeFromAutocomplete
 * @param {any} unitPrice The unit price of the entity value (calling object).
 * @param {number} taxRate The tax rate (0.11 = 11%).
 * @return {number[][]}  Product unit price with tax rate applied.
 */
function applySalesTax(unitPrice, taxRate): number[][] {
  const result: number = unitPrice.basicValue * taxRate + unitPrice.basicValue;
  return [[result]];
}
```

## Add a function to a basic value type

You can add functions to the basic value types of `Boolean`, `double`, and `string`. The process is the same as for entity values. Describe the function with JSON as a lambda method. The following code sample shows how to create a double basic value with a function `AddValue()` that adds a value `x` to the basic value.

```javascript
const lambdaAddValue = {
  type: Excel.CellValueType.lambda,
  body: "Contoso.addValue(callingObject, x)",
  callingObjectAsFirstArgument: true,
  shortDescription: "AddValue",
  longDescription: "Adds x to the object's value",
  example: "=A1.AddValue(5)",
  arguments: [
    {
      name: "callingObject"
    },
    {
      name: "x"
    }
  ]
};

async function createNumberProperties() {
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
            AddValue: lambdaAddValue
          }
        }
      ]
    ];
    await context.sync();
  });
}
```

The following code sample shows the implementation of the `AddValue` custom function.

```javascript
/**
 * Returns the sum of x plus the object's value.
 * @customfunction
 * @excludeFromAutocomplete
 * @param {any} callingObject The basic type being operated on (calling object).
 * @param {number} x Value to add to the calling object's value.
 * @returns {number} x + the calling object's value.
 */
function AddValue(callingObject: any, x: number): number {
  const result: number = callingObject.basicValue;
  return result;
}
```

## Optional arguments

You can specify optional arguments in the JSON using the `LambdaCellValue` type. All optional arguments must be at the end of the function.

The following code sample shows a lambda function named `generateRandomRange` that generates a range of random values. It accepts arguments to specify the number of `rows` and `columns`, and the `min` and `max` of the random values. The `min` and `max` arguments are specified as optional by setting the `optional` property to true.

```javascript
const lambdaRandomRangeWithOptionalArguments = {
  type: Excel.CellValueType.lambda,
  body: "Contoso.generateRandomRange(rows,columns,min,max)",
  callingObjectAsFirstArgument: false, //Set to true to get the calling object
  shortDescription: "GenerateRandomRange",
  longDescription: "Generates a random range of numbers",
  example: "=A1.GenerateRandomRange(2,5,1,10)",
  arguments: [
    {
      name: "rows"
    },
    {
      name: "columns"
    },
    {
      name: "min",
      optional: true
    },
    {
      name: "max",
      optional: true
    }
  ]
};
```

The following code sample shows the custom function named `generateRandomRange` called by the lambda function. It returns a dynamic array of random values matching the number of `rows` and `columns` specified. The `min` and `max` arguments are optional, and if not specified will default to `1` and `10`.

```javascript
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

As the user enters the function in Excel, autocomplete shows the properties of the function, and indicates optional arguments by surrounding them in brackets []. The following screenshot shows an example of the autocomplete for the `generateRandomRange` function.

## Error handling

If your lambda method is using an Excel formula that doesn't call a custom function, any errors will be handled by Excel. The errors returned are the same as if you typed the formula into a cell and ran calculation.
If your lambda method calls a custom function from the formula in the `body` property, your custom function should throw an error if one occurs.

## Differences from the LAMDA function in Excel

Lambda methods on entity values are similar in design to the [LAMDA function](https://support.microsoft.com/office/lambda-function-bd212d27-1cd1-4321-a34a-ccbf254b8b67) in Excel. They both accept arguments, perform a calculation, and return a result. But there are important differences.
-	The LAMDA function in Excel can be added to the Name Manager. Lambda methods on entity values can’t appear in the Name Manager.
-	The LAMDA function in Excel can be entered directly into a cell formula. Lambda methods on entity values can only be called using a dot operator from the entity (`entity.function`).

## Support details

Lambda methods are supported in all custom function types, such as [volatile](custom-functions-volatile.md), and [streaming](custom-functions-web-reqs.md) custom functions. Also, all custom function return types are supported ([matrix, scalar, and error](custom-functions-json-autogeneration.md))
A linked entity can’t have a custom function that combines both a lambda method and a data provider. Be sure when developing linked entities to keep these types of custom functions separate.

## See also

- [Create custom functions in Excel](custom-functions-overview.md)
- [Excel.CellValue](/javascript/api/excel/excel.cellvalue)
- [LinkedEntityCellValue](/javascript/api/excel/excel.linkedentitycellvalue)
- [DoubleCellValue](/javascript/api/excel/excel.doublecellvalue)
- [StringCellValue](/javascript/api/excel/excel.stringcellvalue)
- [BooleanCellValue](/javascript/api/excel/excel.booleancellvalue)
