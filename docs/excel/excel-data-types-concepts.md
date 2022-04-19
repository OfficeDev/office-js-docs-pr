---
title: Excel JavaScript API data types core concepts
description: Learn the core concepts for using Excel data types in your Office Add-in.
ms.date: 04/19/2022
ms.topic: conceptual
ms.prod: excel
ms.custom: scenarios:getting-started
ms.localizationpriority: high
---

# Excel data types core concepts (preview)

> [!NOTE]
> Data types APIs are currently only available in public preview. Preview APIs are subject to change and are not intended for use in a production environment. We recommend that you try them out in test and development environments only. Do not use preview APIs in a production environment or within business-critical documents.
>
> To use preview APIs:
>
> - You must reference the **beta** library on the content delivery network (CDN) (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js). The [type definition file](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) for TypeScript compilation and IntelliSense is found at the CDN and [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts). You can install these types with `npm install --save-dev @types/office-js-preview`. For additional information, see the [@microsoft/office-js](https://www.npmjs.com/package/@microsoft/office-js) NPM package readme.
> - You may need to join the [Office Insider program](https://insider.office.com) for access to more recent Office builds.
>
> To try out data types in Office on Windows, you must have an Excel build number greater than or equal to 16.0.14626.10000. To try out data types in Office on Mac, you must have an Excel build number greater than or equal to 16.55.21102600.

This article describes how to use the [Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) to work with data types. It introduces core concepts that are fundamental to data type development.

## Core concepts

Use the [`Range.valuesAsJson`](/javascript/api/excel/excel.range#excel-excel-range-valuesasjson-member) property to work with data type values. This property is similar to [Range.values](/javascript/api/excel/excel.range#excel-excel-range-values-member), but `Range.values` only returns the four basic types: string, number, boolean, or error values. `Range.valuesAsJson` returns expanded information about the four basic types, and this property can return data types such as formatted number values, entities, and web images.

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

The [CellValueExtraProperties](/javascript/api/excel/excel.cellvalueextraproperties) object is an intersection with rest of the `*CellValue` types. It's not a data type itself. The properties of the `CellValueExtraProperties` object are used with all data types to specify details related to overwriting cell values.

### JSON schema

Each data type uses a JSON metadata schema designed for that type. This defines the [CellValueType](/javascript/api/excel/excel.cellvaluetype) of the data and additional information about the cell, such as `basicValue`, `numberFormat`, or `address`. Each `CellValueType` has properties available according to that type. For example, the `webImage` type includes the [altText](/javascript/api/excel/excel.webimagecellvalue#excel-excel-webimagecellvalue-alttext-member) and [attribution](/javascript/api/excel/excel.webimagecellvalue#excel-excel-webimagecellvalue-attribution-member) properties. The following sections show JSON code samples for the formatted number value, entity value, and web image data types.

The JSON metadata schema for each data type also includes one or more readonly properties that are used when calculations encounter incompatible scenarios, such as a version of Excel that doesn't meet the minimum build number requirement for the data types feature. The property `basicType` is part of the JSON metadata of every data type, and it's always a readonly property. The `basicType` property is used as a fallback when the data type isn't supported or is formatted incorrectly.

## Formatted number values

The [FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue) object enables Excel add-ins to define a `numberFormat` property for a value. Once assigned, this number format travels through calculations with the value and can be returned by functions.

The following JSON code sample shows the complete schema of a formatted number value. The `myDate` formatted number value in the code sample displays as **1/16/1990** in the Excel UI. If the minimum compatibility requirements for the data types feature aren't met, calculations use the `basicValue` in place of the formatted number.

```json
// This is an example of the complete JSON of a formatted number value.
// In this case, the number is formatted as a date.
const myDate: Excel.FormattedNumberCellValue = {
    type: Excel.CellValueType.formattedNumber,
    basicValue: 32889.0,
    basicType: Excel.RangeValueType.double, // A readonly property. Used as a fallback in incompatible scenarios.
    numberFormat: "m/d/yyyy"
};
```

## Entity values

An entity value is a container for data types, similar to an object in object oriented programming. Entities also support arrays as properties of an entity value. The [EntityCellValue](/javascript/api/excel/excel.entitycellvalue) object allows add-ins to define properties such as `type`, `text`, and `properties`. The `properties` property enables the entity value to define and contain additional data types.

The `basicType` and `basicValue` properties define how calculations read this entity data type if the minimum compatibility requirements to use data types aren't met. In that scenario, this entity data type displays as a **#VALUE!** error in the Excel UI.

The following JSON code sample shows the complete schema of an entity value that contains text, an image, a date, and an additional text value.

```json
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
    basicType: Excel.RangeValueType.error, // A readonly property. Used as a fallback in incompatible scenarios.
    basicValue: "#VALUE!" // A readonly property. Used as a fallback in incompatible scenarios.
};
```

## Web image values

The [WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue) object creates the ability to store an image as part of an [entity](#entity-values) or as an independent value in a range. This object offers many properties, including `address`, `altText`, and `relatedImagesAddress`.

The `basicType` and `basicValue` properties define how calculations read the web image data type if the minimum compatibility requirements to use the data types feature aren't met. In that scenario, this web image data type displays as a **#VALUE!** error in the Excel UI.

The following JSON code sample shows the complete schema of a web image.

```json
// This is an example of the complete JSON for a web image.
const myImage: Excel.WebImageCellValue = {
    type: Excel.CellValueType.webImage,
    address: "https://bit.ly/2YGOwtw", 
    basicType: Excel.RangeValueType.error, // A readonly property. Used as a fallback in incompatible scenarios.
    basicValue: "#VALUE!" // A readonly property. Used as a fallback in incompatible scenarios.
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
- [Excel JavaScript API reference](../reference/overview/excel-add-ins-reference-overview.md)
- [Custom functions and data types](custom-functions-data-types-concepts.md)
