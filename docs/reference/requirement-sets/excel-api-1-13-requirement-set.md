---
title: Excel JavaScript API requirement set 1.13
description: 'Details about the ExcelApi 1.13 requirement set.'
ms.date: 06/30/2021
ms.prod: excel
localization_priority: Normal
---

# What's new in Excel JavaScript API 1.13

The ExcelApi 1.13 added a method to insert worksheets into a workbook from a base-64 encoded string and an event to detect workbook activation, and it increased support for formulas in ranges by adding APIs to track changes to formulas and locate a formula's direct dependent cells. It also expanded PivotTable support by adding PivotLayout APIs for alt text, style, and empty cell management.

| Feature area | Description | Relevant objects |
|:--- |:--- |:--- |
| Formula changed events | Track changes to formulas, including the source and type of event that caused a change. | [Worksheet.onFormulaChanged](/javascript/api/excel/excel.worksheet#onFormulaChanged)|
| Formula dependents | Locate the direct dependent cells of a formula. | [Range.getDirectDependents](/javascript/api/excel/excel.range#getDirectDependents__) |
| Insert worksheets | Insert worksheets from another workbook into the current workbook as a base-64 encoded string. | [Workbook.insertWorksheetsFromBase64](/javascript/api/excel/excel.workbook#insertWorksheetsFromBase64_base64File__options_) |
| PivotTable PivotLayout | An expansion of the PivotLayout class, including new support for alt text and empty cell management. | [PivotLayout](/javascript/api/excel/excel.pivotlayout) |

## API list

The following table lists the APIs in Excel JavaScript API requirement set 1.13. To view API reference documentation for all APIs supported by Excel JavaScript API requirement set 1.13 or earlier, see [Excel APIs in requirement set 1.13 or earlier](/javascript/api/excel?view=excel-js-1.13&preserve-view=true).

| Class | Fields | Description |
|:---|:---|:---|


## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel?view=excel-js-1.13&preserve-view=true)
- [Excel JavaScript API requirement sets](excel-api-requirement-sets.md)
