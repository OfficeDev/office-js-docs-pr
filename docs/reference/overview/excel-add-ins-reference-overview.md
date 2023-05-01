---
title: Excel JavaScript API overview
description: Learn more about the Excel JavaScript API.
ms.date: 02/23/2022
ms.service: excel
ms.localizationpriority: high
---

# Excel JavaScript API overview

An Excel add-in interacts with objects in Excel by using the Office JavaScript API, which includes two JavaScript object models:

* **Excel JavaScript API**: These are the [application-specific APIs](../../develop/application-specific-api-model.md) for Excel. Introduced with Office 2016, the [Excel JavaScript API](/javascript/api/excel) provides strongly-typed objects that you can use to access worksheets, ranges, tables, charts, and more.

* **Common APIs**: Introduced with Office 2013, the [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.

This section of the documentation focuses on the Excel JavaScript API, which you'll use to develop the majority of functionality in add-ins that target Excel on the web or Excel 2016 or later. For information about the Common API, see [Common JavaScript API object model](../../develop/office-javascript-api-object-model.md).

## Learn object model concepts

See [Excel JavaScript object model in Office Add-ins](../../excel/excel-add-ins-core-concepts.md) for information about important object model concepts.

For hands-on experience using the Excel JavaScript API to access objects in Excel, complete the [Excel add-in tutorial](../../tutorials/excel-tutorial.md).

## Learn API capabilities

Each major Excel API feature has an article or set of articles exploring what that feature can do and the relevant object model.

* [Charts](../../excel/excel-add-ins-charts.md)
* [Comments](../../excel/excel-add-ins-comments.md)
* [Conditional formatting](../../excel/excel-add-ins-conditional-formatting.md)
* [Custom functions](../../excel/custom-functions-overview.md)
* [Data validation](../../excel/excel-add-ins-data-validation.md)
* [Data types](../../excel/excel-data-types-overview.md)
* [Events](../../excel/excel-add-ins-events.md)
* [PivotTables](../../excel/excel-add-ins-pivottables.md)
* [Ranges](../../excel/excel-add-ins-ranges-get.md) and [Cells](../../excel/excel-add-ins-cells.md)
* [RangeAreas (Multiple ranges)](../../excel/excel-add-ins-multiple-ranges.md)
* [Shapes](../../excel/excel-add-ins-shapes.md)
* [Tables](../../excel/excel-add-ins-tables.md)
* [Workbooks and Application-level APIs](../../excel/excel-add-ins-workbooks.md)
* [Worksheets](../../excel/excel-add-ins-worksheets.md)

For detailed information about the Excel JavaScript API object model, see the [Excel JavaScript API reference documentation](/javascript/api/excel).

## Try out code samples in Script Lab

Use [Script Lab](../../overview/explore-with-script-lab.md) to get started quickly with a collection of built-in samples that show how to complete tasks with the API. You can run the samples in Script Lab to instantly see the result in the task pane or worksheet, examine the samples to learn how the API works, and even use samples to prototype your own add-in.

## See also

* [Excel add-ins documentation](../../excel/index.yml)
* [Excel add-ins overview](../../excel/excel-add-ins-overview.md)
* [Excel JavaScript API reference](/javascript/api/excel)
* [Office client application and platform availability for Office Add-ins](/javascript/api/requirement-sets)
* [Using the application-specific API model](../../develop/application-specific-api-model.md)
