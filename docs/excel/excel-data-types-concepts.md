---
title: Excel JavaScript API data types core concepts
description: Learn the core concepts for using Excel data types in your Office Add-in.
ms.date: 10/10/2022
ms.topic: conceptual
ms.prod: excel
ms.custom: scenarios:getting-started
ms.localizationpriority: high
---

# Excel data types core concepts

This article describes how to use the [Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) to work with data types. It introduces core concepts that are fundamental to data type development.

## The `valuesAsJson` property

The `valuesAsJson` property (or the singular `valueAsJson` for [NamedItem](/javascript/api/excel/excel.nameditem)) is integral to creating data types in Excel. This property is an expansion of `values` properties, such as [Range.values](/javascript/api/excel/excel.range#excel-excel-range-values-member). Both the `values` and `valuesAsJson` properties are used to access the value in a cell, but the `values` property only returns one of the four basic types: string, number, boolean, or error (as a string). In contrast, `valuesAsJson` returns expanded information about the four basic types, and this property can return data types such as formatted number values, entities, and web images.

The following objects offer the `valuesAsJson` property.

- [NamedItem](/javascript/api/excel/excel.nameditem) (as `valueAsJson`)
- [NamedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)
- [Range](/javascript/api/excel/excel.range)
- [RangeView](/javascript/api/excel/excel.rangeview)
- [TableColumn](/javascript/api/excel/excel.tablecolumn)
- [TableRow](/javascript/api/excel/excel.tablerow)

> [!NOTE]
> Some cell values change based on a user's locale. The `valuesAsJsonLocal` property offers localization support and is available on all the same objects as `valuesAsJson`.

## Cell values

The `valuesAsJson` property returns a [CellValue](/javascript/api/excel/excel.cellvalue) type alias, which is a [union](https://www.typescriptlang.org/docs/handbook/2/everyday-types.html#union-types) of the following data types.

- [ArrayCellValue](/javascript/api/excel/excel.arraycellvalue)
- [BooleanCellValue](/javascript/api/excel/excel.booleancellvalue)
- [DoubleCellValue](/javascript/api/excel/excel.doublecellvalue)
- [EmptyCellValue](/javascript/api/excel/excel.emptycellvalue)
- [EntityCellValue](/javascript/api/excel/excel.entitycellvalue)
- [ErrorCellValue](/javascript/api/excel/excel.errorcellvalue)
- [FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue)
- [LinkedEntityCellValue](/javascript/api/excel/excel.linkedentitycellvalue)
- [ReferenceCellValue](/javascript/api/excel/excel.referencecellvalue)
- [StringCellValue](/javascript/api/excel/excel.stringcellvalue)
- [ValueTypeNotAvailableCellValue](/javascript/api/excel/excel.valuetypenotavailablecellvalue)
- [WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue)

The `CellValue` type alias also returns the [CellValueExtraProperties](/javascript/api/excel/excel.cellvalueextraproperties) object, which is an [intersection](https://www.typescriptlang.org/docs/handbook/2/objects.html#intersection-types) with the rest of the `*CellValue` types. It's not a data type itself. The properties of the `CellValueExtraProperties` object are used with all data types to specify details related to overwriting cell values.

### JSON schema

Each cell value type returned by `valuesAsJson` uses a JSON metadata schema designed for that type. Along with additional properties unique to each data type, these JSON metadata schemas all have the `type`, `basicType`, and `basicValue` properties in common.

The `type` defines the [CellValueType](/javascript/api/excel/excel.cellvaluetype) of the data. The `basicType` is always read-only and is used as a fallback when the data type isn't supported or is formatted incorrectly. The `basicValue` matches the value that would be returned by the `values` property. The `basicValue` is used as a fallback when calculations encounter incompatible scenarios, such as an older version of Excel that doesn't support the data types feature. The `basicValue` is read-only for `ArrayCellValue`, `EntityCellValue`, `LinkedEntityCellValue`, and `WebImageCellValue` data types.

In addition to the three fields that all data types share, the JSON metadata schema for each `*CellValue` has properties available according to that type. For example, the [WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue) type includes the `altText` and `attribution` properties, while the [EntityCellValue](/javascript/api/excel/excel.entitycellvalue) type offers the `properties` and `text` fields.

The following sections show JSON code samples for the formatted number value, entity value, and web image data types.

## Formatted number values

The [FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue) object enables Excel add-ins to define a `numberFormat` property for a value. Once assigned, this number format travels through calculations with the value and can be returned by functions.

The following JSON code sample shows the complete schema of a formatted number value. The `myDate` formatted number value in the code sample displays as **1/16/1990** in the Excel UI. If the minimum compatibility requirements for the data types feature aren't met, calculations use the `basicValue` in place of the formatted number.

```TypeScript
// This is an example of the complete JSON of a formatted number value.
// In this case, the number is formatted as a date.
const myDate: Excel.FormattedNumberCellValue = {
    type: Excel.CellValueType.formattedNumber,
    basicValue: 32889.0,
    basicType: Excel.RangeValueType.double, // A read-only property. Used as a fallback in incompatible scenarios.
    numberFormat: "m/d/yyyy"
};
```

## Entity values

An entity value is a container for data types, similar to an object in object-oriented programming. Entities also support arrays as properties of an entity value. The [EntityCellValue](/javascript/api/excel/excel.entitycellvalue) object allows add-ins to define properties such as `type`, `text`, and `properties`. The `properties` property enables the entity value to define and contain additional data types.

The `basicType` and `basicValue` properties define how calculations read this entity data type if the minimum compatibility requirements to use data types aren't met. In that scenario, this entity data type displays as a **#VALUE!** error in the Excel UI.

The following JSON code sample shows the complete schema of an entity value that contains text, an image, a date, and an additional text value.

```TypeScript
// This is an example of the complete JSON for an entity value.
// The entity contains text and properties which contain an image, a date, and another text value.
const myEntity: Excel.EntityCellValue = {
    type: Excel.CellValueType.entity,
    text: "A llama",
    properties: {
        image: myImage,
        "start date": myDate,
        "quote": {
            type: Excel.CellValueType.string,
            basicValue: "I love llamas."
        }
    }, 
    basicType: Excel.RangeValueType.error, // A read-only property. Used as a fallback in incompatible scenarios.
    basicValue: "#VALUE!" // A read-only property. Used as a fallback in incompatible scenarios.
};
```

Entity values also offer a `layouts` property that creates a card for the entity. The card displays as a modal window in the Excel UI and can display additional information contained within the entity value, beyond what's visible in the cell. To learn more, see [Use cards with entity value data types](excel-data-types-entity-card.md).

### Linked entities

Linked entity values, or [LinkedEntityCellValue](/javascript/api/excel/excel.linkedentitycellvalue) objects, are a type of entity value. These objects integrate data provided by an external service and can display this data as an [entity card](excel-data-types-entity-card.md), like regular entity values. The [Stocks and Geography data types](https://support.microsoft.com/office/excel-data-types-stocks-and-geography-61a33056-9935-484f-8ac8-f1a89e210877) available via the Excel UI are linked entity values.

## Web image values

The [WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue) object creates the ability to store an image as part of an [entity](#entity-values) or as an independent value in a range. This object offers many properties, including `address`, `altText`, and `relatedImagesAddress`.

The `basicType` and `basicValue` properties define how calculations read the web image data type if the minimum compatibility requirements to use the data types feature aren't met. In that scenario, this web image data type displays as a **#VALUE!** error in the Excel UI.

The following JSON code sample shows the complete schema of a web image.

```TypeScript
// This is an example of the complete JSON for a web image.
const myImage: Excel.WebImageCellValue = {
    type: Excel.CellValueType.webImage,
    address: "https://bit.ly/2YGOwtw", 
    basicType: Excel.RangeValueType.error, // A read-only property. Used as a fallback in incompatible scenarios.
    basicValue: "#VALUE!" // A read-only property. Used as a fallback in incompatible scenarios.
};
```

## Improved error support

The data types APIs expose existing Excel UI errors as objects. Now that these errors are accessible as objects, add-ins can define or retrieve properties such as `type`, `errorType`, and `errorSubType`.

The following is a list of all the error objects with expanded support through data types.

- [BlockedErrorCellValue](/javascript/api/excel/excel.blockederrorcellvalue)
- [BusyErrorCellValue](/javascript/api/excel/excel.busyerrorcellvalue)
- [CalcErrorCellValue](/javascript/api/excel/excel.calcerrorcellvalue)
- [ConnectErrorCellValue](/javascript/api/excel/excel.connecterrorcellvalue)
- [Div0ErrorCellValue](/javascript/api/excel/excel.div0errorcellvalue)
- [FieldErrorCellValue](/javascript/api/excel/excel.fielderrorcellvalue)
- [GettingDataErrorCellValue](/javascript/api/excel/excel.gettingdataerrorcellvalue)
- [NotAvailableErrorCellValue](/javascript/api/excel/excel.notavailableerrorcellvalue)
- [NameErrorCellValue](/javascript/api/excel/excel.nameerrorcellvalue)
- [NullErrorCellValue](/javascript/api/excel/excel.nullerrorcellvalue)
- [NumErrorCellValue](/javascript/api/excel/excel.numerrorcellvalue)
- [RefErrorCellValue](/javascript/api/excel/excel.referrorcellvalue)
- [SpillErrorCellValue](/javascript/api/excel/excel.spillerrorcellvalue)
- [ValueErrorCellValue](/javascript/api/excel/excel.valueerrorcellvalue)

Each of the error objects can access an enum through the `errorSubType` property, and this enum contains additional data about the error. For example, the `BlockedErrorCellValue` error object can access the [BlockedErrorCellValueSubType](/javascript/api/excel/excel.blockederrorcellvaluesubtype) enum. The `BlockedErrorCellValueSubType` enum offers additional data about what caused the error.

## See also

- [Overview of data types in Excel add-ins](excel-data-types-overview.md)
- [Use cards with entity value data types](excel-data-types-entity-card.md)
- [Excel JavaScript API reference](../reference/overview/excel-add-ins-reference-overview.md)
- [Custom functions and data types](custom-functions-data-types-concepts.md)
