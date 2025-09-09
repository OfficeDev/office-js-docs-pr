---
title: Undo capabilities with the Excel JavaScript API
description: Learn how to preserve the undo stack in your Excel add-ins.
ms.date: 09/09/2025
ms.localizationpriority: medium
---

# Undo support with the Excel JavaScript API

Excel add-ins support undo behavior. This preserves both actions performed by Excel JavaScript APIs and actions performed by the user in Excel. These actions are saved in the *undo stack* for an individual user, allowing the user to step back through their actions when desired.

## Undo grouping

The Excel JavaScript API also supports undo grouping. This allows you to group multiple API calls into a single undoable action for your add-in user. For example, if your add-in needs to make several different updates across multiple worksheets in response to a single user command, you can wrap all those updates in a single group. This is done with the `mergeUndoGroup` property provided to the `Excel.run` function.

If an API within the group doesn't offer undo support, the `UndoNotSupported` error is thrown to let you know that the operation can’t be grouped. Your add-in should gracefully handle this error and present a reasonable message to the user.

The following code sample shows how to merge multiple actions with `mergeUndoGroup` set to `true`.

> [!IMPORTANT]
> Ensure that all grouped API calls support undo to avoid errors. See [Unsupported APIs](#unsupported-apis) for more information.

```js
await Excel.run({ mergeUndoGroup: true }, async (context) => { 
    const sheet = context.workbook.worksheets.getActiveWorksheet(); 
    let range = sheet.getRange("A1"); 
    range.values = [["123"]]; 
    
    await context.sync(); 
    
    range = sheet.getRange("B2"); 
    range.values = [["456"]];

    await context.sync(); 
}); 
```

## Unsupported APIs

Most Excel JavaScript APIs do support undo actions. However, see the following table for a list of APIs that do not support undo behavior.

> [!TIP]
> If you call an unsupported API in your add-in, the user’s undo stack is cleared starting from that API call, and a user cannot undo actions past that point.

| API | Supported in Excel on the web | Supported in Excel on Windows and Excel on Mac | Notes |
|:--------------|:------|:--------|:----------|
| `AllowEditRange.address` | No | No | *None* |
| `AllowEditRange.delete` | No | No | *None* |
| `AllowEditRange.pauseProtection` | No | No | *None* |
| `AllowEditRange.setPassword` | No | No | *None* |
| `AllowEditRange.title` | No | No | *None* |
| `AllowEditRangeCollection.add` | No | No | *None* |
| `AllowEditRangeCollection.pauseProtection` | No | No | *None* |
| `Chart.categoryLabelLevel` | No | No | *None* |
| `Chart.seriesNameLevel` | No | No | *None* |
| `ChartPivotOptions.showAxisFieldButtons` | No | Yes | *None* |
| `ChartPivotOptions.showLegendFieldButtons` | No | Yes | *None* |
| `ChartPivotOptions.showReportFilterFieldButtons` | No | Yes | *None* |
| `ChartPivotOptions.showValueFieldButtons` | No | Yes | *None* |
| `ChartTrendlineLabel.formula` | No | Yes | *None* |
| `DataConnectionCollection.refreshAll` | No | No | *None* |
| `DocumentProperties.author​` | No | Yes | *None* |
| `DocumentProperties.category` | No | Yes | *None* |
| `DocumentProperties.comments` | No | Yes | *None* |
| `DocumentProperties.company` | No | Yes | *None* |
| `DocumentProperties.keywords` | No | Yes | *None* |
| `DocumentProperties.manager` | No | Yes | *None* |
| `DocumentProperties.revisionNumber` | No | Yes | *None* |
| `DocumentProperties.subject` | No | Yes | *None* |
| `DocumentProperties.title` | No | Yes | *None* |
| `LinkedWorkbook.refresh` | No | No | *None* |
| `LinkedWorkbookCollection.refreshAll` | No | No | *None* |
| `NamedItem.comment` | No | Yes | *None* |
| `PivotTableStyle.delete` | No | Yes | API does **not** support co-authoring undo in Excel on Windows and Mac. |
| `PivotTableStyle.duplicate` | No | Yes | *None* |
| `PivotTableStyle.name` | No | Yes | *None* |
| `PivotTableStyleCollection.add` | No | Yes | API does **not** support co-authoring undo in Excel on Windows and Mac. |
| `PivotTableStyleCollection.setDefault` | No | Yes | API does **not** support co-authoring undo in Excel on Windows and Mac. |
| `Query.delete` | No | Yes | API supports undo in Excel on Windows and Mac but doesn't support redo. |
| `Query.refresh` | No | Yes | API supports undo Excel on Windows and Mac but doesn't support redo. |
| `QueryCollection.refreshAll` | No | Yes | API supports undo Excel on Windows and Mac but doesn't support redo. |
| `Slicer.name` | No | Yes | *None* |
| `Slicer.nameInFormula` | No | Yes | *None* |
| `SlicerStyle.delete` | No | Yes | API does **not** support co-authoring undo in Excel on Windows and Mac. |
| `SlicerStyle.duplicate` | No | Yes | *None* |
| `SlicerStyle.name` | No | Yes | *None* |
| `SlicerStyleCollection.add` | No | Yes | API does **not** support co-authoring undo in Excel on Windows and Mac. |
| `SlicerStyleCollection.setDefault` | No | Yes | API does **not** support co-authoring undo in Excel on Windows and Mac. |
| `Style.addIndent` | No | Yes | *None* |
| `Style.autoIndent` | No | Yes | *None* |
| `Style.formulaHidden` | No | Yes | *None* |
| `Style.horizontalAlignment` | No | Yes | *None* |
| `Style.includeAlignment` | No | Yes | *None* |
| `Style.includeBorder` | No | Yes | *None* |
| `Style.includeFont` | No | Yes | *None* |
| `Style.includeNumber` | No | Yes | *None* |
| `Style.includePatterns` | No | Yes | *None* |
| `Style.includeProtection` | No | Yes | *None* |
| `Style.indentLevel` | No | Yes | *None* |
| `Style.locked` | No | Yes | *None* |
| `Style.numberFormat` | No | Yes | *None* |
| `Style.numberFormatLocal` | No | Yes | *None* |
| `Style.orientation` | No | Yes | *None* |
| `Style.readingOrder` | No | Yes | *None* |
| `Style.shrinkToFit` | No | Yes | *None* |
| `Style.textOrientation` | No | Yes | *None* |
| `Style.verticalAlignment` | No | Yes | *None* |
| `Style.wrapText` | No | Yes | *None* |
| `TableStyle.delete` | No | Yes | API does **not** support co-authoring undo in Excel on Windows and Mac. |
| `TableStyle.duplicate` | No | Yes | *None* |
| `TableStyle.name` | No | Yes | *None* |
| `TableStyleCollection.add` | No | Yes | API does **not** support co-authoring undo in Excel on Windows and Mac. |
| `TableStyleCollection.setDefault` | No | Yes | API does **not** support co-authoring undo in Excel on Windows and Mac. |
| `TimelineStyle.delete` | No | Yes | API does **not** support co-authoring undo in Excel on Windows and Mac. |
| `TimelineStyle.duplicate` | No | Yes | *None* |
| `TimelineStyle.name` | No | Yes | *None* |
| `TimelineStyleCollection.add` | No | Yes | API does **not** support co-authoring undo in Excel on Windows and Mac. |
| `TimelineStyleCollection.setDefault` | No | Yes | API does **not** support co-authoring undo in Excel on Windows and Mac. |
| `Workbook.close` | No | No | *None* |
| `Workbook.insertWorksheetsFromBase64` | No | No | *None* |
| `Workbook.save` | No | No | *None* |
| `WorkbookProtection.protect` | No | No | *None* |
| `WorkbookProtection.unprotect` | No | No | *None* |
| `Worksheet.copy` | No | No | *None* |
| `Worksheet.delete` | No | No | *None* |
| `Worksheet.name` | Yes | No | *None* |
| `Worksheet.standardWidth` | No | Yes | *None* |
| `Worksheet.position` | Yes | No | *None* |
| `Worksheet.visibility​` | Yes | No | *None* |
| `WorksheetCollection.addFromBase64` | No | No | *None* |
| `WorksheetProtection.pauseProtection` | No | No | *None* |
| `WorksheetProtection.protect` | No | No | *None* |
| `WorksheetProtection.resumeProtection` | No | No | *None* |
| `WorksheetProtection.setPassword` | No | No | *None* |
| `WorksheetProtection.unprotect` | No | No | *None* |
| `WorksheetProtection.updateOptions` | No | No | *None* |

## See also

- [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md)