---
title: Synchronous custom functions
description: Use synchronous custom functions to support evaluation and conditional format processes in Excel.
ms.date: 12/09/2025
ms.localizationpriority: medium
---

# Synchronous custom functions

Synchronous custom functions enable the evaluation and conditional format processes to run in Excel simultaneously with the custom function. Enable synchronous support when your custom function needs to run in tandem with any of the Excel processes listed in this article. Custom functions that don't support synchronous operations return an error such as `#CALC!` or `#VALUE!` when they run at the same time as these Excel processes.

> [!WARNING]
> Synchronous custom functions don't support write operations with Office JavaScript APIs, such as using `Range.values` to set a cell value. Calling a write operation in a synchronous custom function may cause Excel to freeze.

## Excel processes supported by synchronous custom functions

The following actions and processes work with synchronous custom functions.

### Supported evaluation actions

- Selecting Formulas and then Evaluate Formula in Excel.
- Selecting Formulas and then Insert Function in Excel.
- In cell edit mode, selecting part of a formula and using F9 to see partial calculation results.
- The `Application.Calculate` method in VBA.

### Supported conditional format actions

The following list applies to conditional formatting actions triggered by both the Excel UI and Office JavaScript APIs.

- Create a new rule.
- Edit rules.
- Delete rules.
- Reorder rules.
- Change the “Applies to” range.
- Toggle “Stop if true” on or off.
- Clear all rules.
- Cut and paste cells that contain conditional formatting.
- Copy and paste cells that contain conditional formatting.

> [!NOTE]
> When a synchronous custom function takes a significant amount of time to complete, Excel might temporarily block the user interface while waiting for the result. To avoid prolonged interruptions, users can cancel the function execution at any time by using the <kbd>Esc</kbd> key or selecting anywhere outside the cell or dialog.

## Enable synchronous support in your add-in

To support synchronous scenarios in your add-in, take the following steps.

- If you [autogenerate JSON metadata](custom-functions-json-autogeneration.md), use the `@supportSync` JSDoc tag.
- If you [manually create JSON metadata](custom-functions-json.md), use the `"supportSync": true` setting in the `"options"` object of your **functions.json** file.
- If the function uses `Excel.RequestContext`, call the `setInvocation` method of `Excel.RequestContext` and pass in the `CustomFunctions.Invocation` object. For an example, see the [code sample](#code-sample).

> [!IMPORTANT]
> Synchronous custom functions can't be **streaming** or **volatile** functions. If you use the `@supportSync` tag with `@volatile` or `@streaming` tags, Excel ignores the synchronous support. Volatile or streaming support takes precedence.

### Code sample

The following code sample shows how to create a synchronous custom function.

```typescript
/** 
 * A synchronous custom function that takes a cell address and returns the value of that cell.
 * @customfunction
 * @supportSync
 * @param {string} address The address of the cell from which to retrieve the value.
 * @param {CustomFunctions.Invocation} invocation Invocation object.
 * @returns The value of the cell at the input address.
 */ 
export async function getCellValue(address, invocation) {
  const context = new Excel.RequestContext();
  context.setInvocation(invocation); // The `invocation` object must be passed in the `setInvocation` method for synchronous functions.

  const range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
  range.load("values");

  await context.sync(); 
  return range.values[0][0];
}
```

## See also

- [Autogenerate JSON metadata for custom functions](custom-functions-json-autogeneration.md)
- [Manually create JSON metadata for custom functions](custom-functions-json.md)
