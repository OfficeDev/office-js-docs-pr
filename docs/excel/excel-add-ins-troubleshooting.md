---
title: Troubleshooting Excel Add-ins
description: 'Learn how to troubleshoot development errors in Excel Add-ins.'
ms.date: 09/08/2020
localization_priority: Normal
---

# Troubleshooting Excel Add-ins

## API limitations when the active workbook switches

Add-ins for Excel are intended to operate on a single workbook at a time. Errors can arise when a workbook that is separate from the one running the add-in gains focus. This only happens when particular methods are in the process of being called when the focus changes.

The following APIs are affected by this workbook switch:

|Excel JavaScript API | Error thrown |
|--|--|
| `Chart.activate` | GeneralException |
| `Range.select` | GeneralException |
| `Table.clearFilters` | GeneralException |
| `Workbook.getActiveCell`  | InvalidSelection|
| `Workbook.getSelectedRange` | InvalidSelection|
| `Workbook.getSelectedRanges`  | InvalidSelection|
| `Worksheet.activate` | GeneralException |
| `Worksheet.delete`  | InvalidSelection|
| `Worksheet.gridlines` | GeneralException |
| `Worksheet.showHeadings` | GeneralException |
| `WorksheetCollection.add` | GeneralException |
| `WorksheetFreezePanes.freezeAt` | GeneralException |
| `WorksheetFreezePanes.freezeColumns` | GeneralException |
| `WorksheetFreezePanes.freezeRows` | GeneralException |
| `WorksheetFreezePanes.getLocationOrNullObject`| GeneralException |
| `WorksheetFreezePanes.unfreeze` | GeneralException |

> [!NOTE]
> This only applies to multiple Excel workbooks open on Windows or Mac.

## Coauthoring

See [Coauthoring in Excel add-ins](co-authoring-in-excel-add-ins.md) for patterns to use with events in a coauthoring environment. The article also discusses potential merge conflicts when using certain APIs, such as [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add-index--values-).

## See also

- [Troubleshoot development errors with Office Add-ins](../testing/troubleshoot-development-errors.md)
- [Troubleshoot user errors with Office Add-ins](../testing/testing-and-troubleshooting.md)
