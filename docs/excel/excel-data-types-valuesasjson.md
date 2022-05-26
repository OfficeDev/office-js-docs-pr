---
title: Excel JavaScript API data types valuesAsJson
description: Learn how to build data types with the valuesAsJson property in your Excel add-in.
ms.date: 05/24/2022
ms.topic: conceptual
ms.prod: excel
ms.localizationpriority: medium
---

- valuesAsJson and a list of the objects that offer this property (Range, TableRow, TableColumn, etc.)
- purpose of valuesAsJsonLocal
- An explanation of the CellValue type alias as a union of multiple types
- CellValueExtraProperties as an intersection with the rest of the CellValue types

# `valuesAsJson`

The `valuesAsJson` property is integral to creating data types in Excel. This property is an expansion of the `values` property available on multiple Excel JavaScript API objects. Both of these properties are used to access the value in a cell. The `values` property only returns one of the four basic types: string, number, boolean, or error values. In contrast, `valuesAsJson` returns expanded information about the four basic types, and this property can return data types such as formatted number values, entities, and web images.

The following objects offer the `valuesAsJson` property.

- Range
- TableRow
- TableColumn

> [!NOTE]
> Some cell values change based on a user's locale. The `valuesAsJsonLocal` property offers localization support and is available on all the same objects as `valuesAsJson`.

## Cell values

The `valuesAsJson` property returns a [CellValue](/javascript/api/excel/excel.cellvalue) type alias, which is a [union](https://www.typescriptlang.org/docs/handbook/2/everyday-types.html#union-types) of the following data types.

- [ArrayCellValue](/javascript/api/excel/excel.arraycellvalue)
- [BooleanCellValue](/javascript/api/excel/excel.booleancellvalue)
- [DoubleCellValue](/javascript/api/excel/excel.doublecellvalue)
- [EntityCellValue](/javascript/api/excel/excel.entitycellvalue)
- [EmptyCellValue](/javascript/api/excel/excel.emptycellvalue)
- [ErrorCellValue](/javascript/api/excel/excel.errorcellvalue)
- [FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue)
- [LinkedEntityCellValue](/javascript/api/excel/excel.linkedentitycellvalue)
- [ReferenceCellValue](/javascript/api/excel/excel.referencecellvalue)
- [StringCellValue](/javascript/api/excel/excel.stringcellvalue)
- [ValueTypeNotAvailableCellValue](/javascript/api/excel/excel.valuetypenotavailablecellvalue)
- [WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue)

All cell values returned by `valuesAsJson` have three fields in common, `type`, `basicType`, and `basicValue`. The `type` field allows us to inspect any cell value, even a cell value that is one of the four basic types (string, number, boolean, or error values), as though it were a complex type. Because Excel has a long history of processing simple data as only text, numbers, booleans, or errors, all cell values also have the `basicType` field. The `basicType` field only supports string, number, boolean, and error values. This approach, offering basic types, complex types, and complex types that can be used as simple types, aligns with most programming languages.

## Programming approach to the cell value schemas

Cell value interfaces are unrelated by [inheritance](https://en.wikipedia.org/wiki/Inheritance_(object-oriented_programming)), but all cell values have a `type` field which allows for [type inference](https://en.wikipedia.org/wiki/Type_inference).

```TypeScript
value = range.valuesAsJson[0][0];

switch(value.type) {
    case "Entity":
        console.log(value.text); // no error because TS knows that only an EntityCellValue has a type property of "Entity"
    break;
    case "WebImage":
        // here's the alternative - you have to explicitly specify the type to access its fields
        console.log((value as Excel.WebImageCellValue).address);
    break;
}
```

All cell values have three fields: `type`, `basicType`, and `basicValue`. The `type` field allows us to inspect any value as though it were a complex type. However, Excel has a long history of processing simple data as only text, numbers, booleans, or errors. To support simple data scenarios, we use `basicType`. This is also aligned with most programming languages which have basic types, complex types, and complex types which can be used a simple types.

Example: If you want to sum a column of numbers you might be tempted to sum the values of all the cells where `type` was "Double". However, Excel also has formatted numbers which have "FormattedNumber" in their `type` property. Instead, check that `basicType` is "Double". Any cell value which wishes to be treated as a number will have a `basicType` of "Double".

Some cell values shouldn't be treated Cell values which shouldn't be treated as a basic type typically have a `basicType` of "Error" and `basicValue` of "#VALUE!".

If all three fields weren't present, or if `valuesAsJson` returned basic values in Excel as basic values in JavaScript, then add-in writers would need much more complex checks to identify the type of a value in a cell.

The `valuesAsJson` property has a JSON schema so that update operations can be simulated with read and create actions. This allows users to write code in a style consistent with JavaScript while being consistent with Office.js's architecture of batched calls. Alternative would be to have a specialized schema for updates or tracking cell values themselves on the front-end. The former has enormous learning requirements for users and the latter would have poor performance and be difficult to implement.

```TypeScript
// Read
const range = context.workbook.worksheets.getActiveWorksheet().getRange("A1:B2");
var values = range.valuesAsJson;
await context.sync();

// "Update" B2 (mutates the serialized data, not the actual values in Excel)
values[1][1].text = "Updated Text";
// other updates...

// Create (this overwrites the entire range - but values which weren't changed during the "Update" will be identical afterwards and so this will look like an Update operation)
range.valuesAsJson = values;
await context.sync();
```

The `valuesAsJson` property treats data literally. This means that strings like **"=1+2"** and **"MARCH1"** become the text of a cell. That is different from `values`, which treats all inputs as though they had been typed into Excel's formula bar and so parses them. This means that from `values`, **"=1+2"** becomes a formula whose result is 3 and **"MARCH1"** becomes a number whose cell has date formatting applied. While this is a departure from existing behavior, it offers lossless serialization of cell values and enables the elegant "Update" scenario above.

This may be a controversial point and is one of the reasons why `valuesAsJson` is still in preview. It creates a situation where if data is re-entered (select a cell, press F2 to edit, then ENTER to commit), the value will change without warning.

Treating data literally means that data entry is slightly faster through `valuesAsJson` than `values`, because parsing is bypassed.

The `valuesAsJson` property will always return JSON objects because `type`, `basicType`, and `basicValue` are guaranteed to be present. However, it will accept JavaScript strings (treated as `StringCellValue`), numbers (treated as `DoubleCellValue`), and booleans (treated as `BooleanCellValue`).

IntelliSense will not show that JS basic values are eligible for input. This is because the minimum version of TypeScript required by Office.js doesn't support variant accessors. If we had added the basic values to the type of the `valuesAsJson` property, type inference based on the `type` (see first point) wouldn't work because TypeScript would think that the cell value might be a string or number, not an object.

The data which `valuesAsJson` uses to represent a cell value needs to be sufficient to exactly re-recreate the value, we don't want to include unnecessary information because of bandwidth issues and the time it might take to calculate some fields. Our thinking was to include all "non-derived" properties (properties of a cell's value which can't be derived from other properties).

Example: FormattedNumberCellValue doesn't have a `text` property representing the text which would in appear in the cell when the number formatting in the value is applied to its number. In this particular case, `values` can be used to get the text shown in the cell, though the value's number formatting may have been overridden by cell formatting. In the case of a formatted number which is in `ArrayCellValue.elements` or `EntityCellValue.properties`, there's is no way of getting the text which the user will see after formatting has been applied. [When there is customer demand, we expect to add a "queries API" which will allow add-in developers to query for derived properties.]

// The `valuesAsJsonLocal` property exists because some properties in cell values will change or be interpreted different depending on the user's locale.

The `basicValue` value on every interface which represents an error in a cell is localized. If it weren't, it would be redundant with the `errorValue` property which indicates the type of error in the cell. `errorValue` can be used to write locale-independent code.

## See also

- [Overview of data types in Excel add-ins](excel-data-types-overview.md)
- [Excel data types core concepts](excel-data-types-concepts.md)
- [Use cards with entity value data types](excel-data-types-entity-card.md)
- [Excel JavaScript API reference](../reference/overview/excel-add-ins-reference-overview.md)