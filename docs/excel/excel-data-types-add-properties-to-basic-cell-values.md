---
title: Add properties to basic cell values
description: Add properties to basic cell values.
ms.topic: how-to #Required; leave this attribute/value as-is
ms.date: 05/12/2025
ms.localizationpriority: medium
---

# Add properties to basic cell values

Add properties to basic cell values in Excel to associate additional information with the values. Similar to [entity values](excel-data-types-linked-entity-cell-values.md), you can add properties to the **string**, **double**, and **Boolean** basic types. Each property is a key/value pair. The following example shows the number 14.67 (a double) that represents a bill with added fields named **Drinks**, **Food**, **Tax**, and **Tip**.

:::image type="content" source="../images/data-type-basic-fields.png" alt-text="Screenshot of the drinks, food, tax, and tip fields shown for the selected cell value.":::

If the user chooses to show the data type card, they'll see the values for the fields.

:::image type="content" source="../images/data-type-basic-data-type-card.png" alt-text="Data type card showing values for drinks, food, tax, and tip properties.":::

Cell value properties can also be used in formulas.

:::image type="content" source="../images/data-type-basic-dot-syntax.png" alt-text="Show user typing 'a1.' and Excel showing a menu with drinks, food, tax, and tip options.":::

## Create a cell value with properties

To create a cell value and add properties to it, use `Range.valuesAsJson` to assign properties. The following code sample shows how to create a new number in cell **A1**. It adds the **Food**, **Drinks**, and additional properties describing a bill in a restaurant. It assigns a JSON description of the properties to `valuesAsJson`.

```typescript
async function createNumberProperties() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange("A1");
    range.valuesAsJson = [
      [
        {
          type: Excel.CellValueType.double,
          basicType: Excel.RangeValueType.double,
          basicValue: 14.67,
          properties: {
            Food: {
              type: Excel.CellValueType.string,
              basicType: Excel.RangeValueType.string,
              basicValue: "Sandwich and fries"
            },
            Drinks: {
              type: Excel.CellValueType.string,
              basicType: Excel.RangeValueType.string,
              basicValue: "Soda"
            },
            Tax: {
              type: Excel.CellValueType.double,
              basicType: Excel.RangeValueType.double,
              basicValue: 5.5
            },
            Tip: {
              type: Excel.CellValueType.double,
              basicType: Excel.RangeValueType.double,
              basicValue: 21
            }
          }
        }
      ]
    ];
    await context.sync();
  });
}
```

> [!NOTE]
> Some cell values change based on a user's locale. The `valuesAsJsonLocal` property offers localization support and is available on all the same objects as `valuesAsJson`.

## Add properties to an existing value

To add properties to an existing value, first get the value from the cell using `valuesAsJson`, then add a properties JSON object to it. The following example shows how to get the number value from cell **A1** and assign a property named **Precision** to it. Note that you should check the type of the value to ensure it's a **string**, **double**, or **Boolean** basic type.

```typescript
async function addPropertyToNumber() {
    await Excel.run(async (context) => {
        let sheet = context.workbook.worksheets.getActiveWorksheet();
        let range = sheet.getRange("A1");
        range.load("valuesAsJson");
        await context.sync();
        let cellValue = range.valuesAsJson[0][0] as any;

        // Only apply this property to a double.
        if (cellValue.basicType === "Double") {
            cellValue.properties = {
                Precision: {
                    type: Excel.CellValueType.double,
                    basicValue: 4
                }
            };
            range.valuesAsJson = [[cellValue]];
            await context.sync();
        }
    });
}
```

## Differences from entity values

Adding properties to **string**, **Boolean**, and **double** basic types is similar to adding properties to entity values. However there are differences.

- Basic types have a non-error fallback so that calculations can operate on them. For example, consider the formula **=SUM(A1:A3)** where **A1** is **1**, **A2** is **2**, and **A3** is **3**. **A1** is a double with properties, while **A2** and **A3** don't have properties. The sum returns the correct result of **6**. The formula wouldn't work if **A1** was an entity value.
- When the value of a basic type is used in a calculation, the properties are excluded in the result. In the previous example of **=SUM(A1:A3)** where A1 is a double with properties, the result of **6** does not have any properties.
- If no icon is specified for a basic type, the cell doesn't show any icon. But if an entity value doesn't specify an icon, it shows a default icon in the cell value.

## Formatted number values

You can apply number formatting to values of type `CellValueType.double`. Use the `numberFormat` property in the JSON schema to specify a number format. The following code sample shows the complete schema of a number value formatted as currency. The formatted number value in the code sample displays as **$24.00** in the Excel UI.

```typescript
// This is an example of the complete JSON of a formatted number value with a property.
// In this case, the number is formatted as currency.
async function createCurrencyValue() {
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange("A1");
        range.valuesAsJson = [
            [
                {
                    type: Excel.CellValueType.double,
                    basicType: Excel.RangeValueType.double,
                    basicValue: 24,
                    numberFormat: "$0.00",
                    properties: {
                        Name: {
                            type: Excel.CellValueType.string,
                            basicValue: "dollar"
                        }
                    }
                }
            ]
        ];
        await context.sync();
    });
}
```

The number formatting is considered the default format. If the user, or other code, applies formatting to a cell containing a formatted number, the applied format overrides the number’s format.

## Card layout

Cell values with properties have a default data type card that the user can view. You can provide a custom card layout to use instead of the default card layout to improve the user experience when viewing properties. To do this, add the **layouts** property to the JSON description.

For more information, see [Use cards with cell value data types](excel-data-types-entity-card.md).

## Nested data types

You can nest data types in a cell value, such as additional entity values, as well as **strings**, **doubles**, and **Booleans**. The following code sample shows how to create a cell value that represents the charge status on a computer battery. It contains a nested entity value that describes the computer properties for power consumption and charging status. The computer entity value also contains a nested string value that describes the computer’s power plan.

```typescript
async function createNumberWithNestedEntity() {
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange("A1");
        range.valuesAsJson = [
            [
                {
                    type: Excel.CellValueType.double,
                    basicType: Excel.RangeValueType.double,
                    layouts: {
                        compact: {
                            icon: "Battery10"
                        }
                    },
                    basicValue: 0.7,
                    numberFormat: "00%",
                    properties: {
                        Computer: {
                            type: Excel.CellValueType.entity,
                            text: "Laptop",
                            properties: {
                                "Power Consumption": {
                                    type: Excel.CellValueType.double,
                                    basicType: Excel.RangeValueType.double,
                                    basicValue: 0.25,
                                    numberFormat: "00%",
                                    layouts: {
                                        compact: {
                                            icon: "Power"
                                        }
                                    },
                                    properties: {
                                        plan: {
                                            type: Excel.CellValueType.string,
                                            basicType: Excel.RangeValueType.string,
                                            basicValue: "Balanced"
                                        }
                                    }
                                },
                                Charging: {
                                    type: Excel.CellValueType.boolean,
                                    basicType: Excel.RangeValueType.boolean,
                                    basicValue: true
                                }
                            }
                        }
                    }
                }
            ]
        ];
        await context.sync();
    });
}
```

The following image shows the number value and the data type card for the nested laptop entity.

:::image type="content" source="../images/data-type-basic-nested-entities.png" alt-text="Cell value in Excel showing battery charge at 70%, and the data type card showing the nested laptop entity with charging and power consumption property values.":::

## Compatibility

On previous versions of Excel that don't support the data types feature, users see a warning of **Unavailable Data Type**. The value still displays in the cell and functions as expected with formulas and other Excel features. If the value is a formatted number, calculations use the `basicValue` in place of the formatted number.

On Excel versions older than Office 2016, the value is shown in the cell with no error and is indistinguishable from a basic value.

## See also

- [Excel JavaScript API data types core concepts](excel-data-types-concepts.md)
