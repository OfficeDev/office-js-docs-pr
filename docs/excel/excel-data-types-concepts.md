---
title: Excel JavaScript API custom data types core concepts
description: 'Learn the core concepts for using Excel custom data types in your Office Add-in.'
ms.date: 10/06/2021
ms.topic: conceptual
ms.prod: excel
ms.custom: scenarios:getting-started
ms.localizationpriority: high
---

# Excel custom data types core concepts (preview)

> [!NOTE]
> Custom data types APIs are currently only available in public preview. [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]
> 

This article describes how to use the [Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) to work with custom data types. It introduces core concepts that are fundamental to custom data type development and provides guidances for performing specific tasks such as reading or writing to a custom data type.

## Core concepts

- The type property
- Range.value and Range.valueAsJSON
- basicType, basicValue

## Read and write a custom data type

### Cell values

- [BooleanCellValue](/javascript/api/excel/excel.booleancellvalue)
- [DoubleCellValue](/javascript/api/excel/excel.doublecellvalue)
- [EmptyCellValue](/javascript/api/excel/excel.emptycellvalue)
- [StringCellValue](/javascript/api/excel/excel.stringcellvalue)
- [ValueTypeNotAvailableCellValue](/javascript/api/excel/excel.valuetypenotavailablecellvalue)

### Formatted number values

- [FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue)

### Web image

- [WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue)

### Rich errors

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

### Entity schema

TBD.

## See also

- [Excel custom data types core concepts](/excel-data-types-concepts.md)
- [Excel JavaScript API reference](../reference/overview/excel-add-ins-reference-overview.md)