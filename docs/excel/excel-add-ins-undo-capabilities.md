---
title: Undo capabilities with the Excel JavaScript API
description: Learn how to preserve the undo stack in your Excel add-ins.
ms.date: 05/15/2025
ms.localizationpriority: medium
---

# Undo support with the Excel JavaScript API

Excel add-ins support undo behavior, preserving both actions performed by Excel JavaScript APIs and actions performed by the user in Excel. These actions are preserved in the *undo stack* for an individual user, allowing the user to step back through their actions when desired.

## Undo grouping

The Excel JavaScript API also supports undo grouping, which allows you to group multiple API calls into a single undoable action for your add-in user. For example, if your add-in needs to make several different updates across multiple worksheets in response to a single user command, you can wrap all those updates in a single group.

If an API within the group doesn't offer undo support, the `UndoNotSupported` error is thrown to let you know that the operation can’t be grouped.

The following code sample shows how to merge multiple actions with `mergeUndoGroup` set to `true`.

> [!IMPORTANT]
> Ensure that all grouped API calls support undo to avoid errors. See [Unsupported APIs](#unsupported-apis) and [Check for undo support](#check-for-undo-support) for more information.

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

Most Excel JavaScript APIs support undo actions. See the following table for a list of APIs that don't support undo behavior.

> [!TIP] If you call an unsupported API in your add-in, the user’s undo stack is cleared starting from that API call, and a user cannot undo actions past that point.

## Check for undo support

Use the`isSetSupported` method of the [OfficeRuntime.ApiInformation](/javascript/api/office-runtime/officeruntime.apiinformation) interface to check for undo support and provide a fallback experience if undo support isn't available. See the following code sample for more information about how to use `isSetSupported`.

```js
    if (Office.context.requirements.isSetSupported("ExcelUndoApiAll", "1.0")) { 
       // Undo is supported. 
    } else { 
       // Undo is not supported.
    } 
```

## See also

- [Excel JavaScript object model in Office Add-ins](excel-add-ins-core-concepts.md)
