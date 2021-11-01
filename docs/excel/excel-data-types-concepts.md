---
title: Excel JavaScript API data types core concepts
description: 'Learn the core concepts for using Excel data types in your Office Add-in.'
ms.date: 10/27/2021
ms.topic: conceptual
ms.prod: excel
ms.custom: scenarios:getting-started
ms.localizationpriority: high
---

# Excel data types core concepts (preview)

> [!NOTE]
> Data types APIs are currently only available in public preview. [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]
> 

This article describes how to use the [Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) to work with data types. It introduces core concepts that are fundamental to data type development.

> [!IMPORTANT]
> Some of the data types concepts described in this article, such as `Range.valueAsJSON` are not yet available in public preview. This article is intended as a conceptual introduction. Concepts described in this article that are not yet in public preview will be released to preview soon.

## Core concepts

Use the `Range.valueAsJSON` property to work with data type values. This property is similar to [Range.values](/javascript/api/excel/excel.range#values), but `Range.values` only returns the four basic types: string, number, boolean, or error values. `Range.valueAsJSON` can return expanded information about the four basic types, and this property can return data types such as formatted number values, entities, and web images.

### JSON schema

Data types use a consistent JSON schema which defines the [CellValueType](/javascript/api/excel/excel.cellvaluetype) of the data and additional information such as `basicValue`, `numberFormat`, or `address`. Each `CellValueType` has properties available according to that type. For example, the `webImage` type includes the [altText](/javascript/api/excel/excel.webimagecellvalue#altText) and [attribution](/javascript/api/excel/excel.webimagecellvalue#attribution) properties. The following sections show JSON code samples for the formatted number value, entity value, and web image data types.

## Formatted number values

The [FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue) object enables Excel add-ins to define a `numberFormat` property for a value. Once assigned, this number format travels through calculations with the value and can be returned by functions.

The following JSON code sample shows a formatted number value. The `myDate` formatted number value in the code sample displays as **1/16/1990** in the Excel UI.

```json
// This is an example of the JSON schema of a formatted number value.
// In this case, the number is formatted as a date.
const myDate = {
    type: Excel.CellValueType.formattedNumber,
    basicValue: 32889.0,
    numberFormat: "m/d/yyyy"
};
```

## Entity values

An entity value is a container for data types, similar to an object in object oriented programming. Entities also support arrays as properties of an entity value.

The following JSON code sample shows an entity value that contains text, an image, a date, and an additional text value.

```json
// This is an example of the JSON for an entity value.
// The entity contains text and properties which contain an image, a date, and another text value.
const myEntity = {
    type: Excel.CellValueType.entity,
    text: "A llama",
    properties: {
        image: myImage,
        "start date": myDate,
        "quote": {
            type: Excel.CellValueType.string,
            basicValue: "I love llamas."
        }
    }
};
```

## Web image values

The [WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue) object creates the ability to store an image as part of an [entity](#entity-values) or as an independent value in a range. This object offers many properties, including `address`, `altText`, and `relatedImagesAddress`.

The following JSON code sample shows how to represent a web image.

```json
// This is an example of the JSON for a web image.
const myImage = {
    type: Excel.CellValueType.webImage,
    address: "https://bit.ly/2YGOwtw"
};
```

## Improved error support

The improved error support included in the data types APIs allows access to the properties contained within errors returned by the Excel UI. The following is a list of all the error objects with expanded support through data types.

- [BlockedErrorCellValue](/javascript/api/excel/excel.blockederrorcellvalue)
- [BusyErrorCellValue](/javascript/api/excel/excel.busyerrorcellvalue)
- [CalcErrorCellValue](/javascript/api/excel/excel.calcerrorcellvalue)
- [ConnectErrorCellValue](/javascript/api/excel/excel.connecterrorcellvalue)
- [Div0ErrorCellValue](/javascript/api/excel/excel.div0errorcellvalue)
- [FieldErrorCellValue](/javascript/api/excel/excel.fielderrorcellvalue)
- [GettingDataErrorCellValue](/javascript/api/excel/excel.gettingdataerrorcellvalue)
- [NaErrorCellValue](/javascript/api/excel/excel.naerrorcellvalue)
- [NameErrorCellValue](/javascript/api/excel/excel.nameerrorcellvalue)
- [NullErrorCellValue](/javascript/api/excel/excel.nullerrorcellvalue)
- [NumErrorCellValue](/javascript/api/excel/excel.numerrorcellvalue)
- [RefErrorCellValue](/javascript/api/excel/excel.referrorcellvalue)
- [SpillErrorCellValue](/javascript/api/excel/excel.spillerrorcellvalue)
- [ValueErrorCellValue](/javascript/api/excel/excel.valueerrorcellvalue)

## See also

- [Excel data types core concepts](/excel-data-types-concepts.md)
- [Excel JavaScript API reference](../reference/overview/excel-add-ins-reference-overview.md)
- [Custom functions and data types overview](/custom-functions-data-types-overview.md)