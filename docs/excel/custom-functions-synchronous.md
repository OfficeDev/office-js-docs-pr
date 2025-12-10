---
title: Synchronous custom functions
description: Use synchronous custom functions to support evaluation and conditional format processes in Excel.
ms.date: 12/09/2025
ms.localizationpriority: medium
---

# Synchronous custom functions

Synchronous custom functions allow concurrent evaluate and conditional format processes in Excel. Without an explicit synchronous setting, custom functions can't run during these Excel processes.

## Excel processes supported by synchronous custom functions

The following actions and processes are supported with synchronous custom functions.

### Evaluate actions

- UI action: Formulas > Evaluate Formula.
- UI action: Formulas > Insert Function.
- UI action: In cell edit mode, selecting part of a formula and using F9 to see partial calculation results.
- VBA API: Application.Calculate

### Conditional format actions

The following list applies to both UI actions and Office JavaScript API actions.

- Create new rule
- Edit rules
- Delete rules
- Reorder rules
- Change “Applies to” range
- Toggle “Stop if True”
- Clear all rules
- Copy/Cut and Paste cells containing conditional formatting

> [!NOTE]
> When a synchronous custom function takes a significant amount of time to complete, Excel may temporarily block the user interface while waiting for the result. To avoid prolonged interruptions, users can cancel the execution at any time by using <kbd>Esc</kbd> or by selecting anywhere outside the cell or dialog.

## Enable synchronous support in your add-in

To support synchronous scenarios in your add-in:

- If you [autogenerate JSON metadata](custom-functions-json-autogeneration.md), use the `@supportSync` JSDoc tag.
- If you [manually create JSON metadata](custom-functions-json.md), use the `"supportSync": true` setting in the `"options"` object of your **functions.json** file.
- If the function uses `Excel.RequestContext`, call the `setInvocation` method of `Excel.RequestContext` and pass in the `CustomFunctions.Invocation` object. For an example, see the [synchronous custom function code sample](#synchronous-custom-function-code-sample).

> [!NOTE]
> Synchronous custom functions cannot be **streaming** or **volatile** functions.

### Synchronous custom function code sample

```typescript
/** 
 * @customfunction
 * @supportSync
 * @param {string} address The address of the cell from which to retrieve the value.
 * @param {CustomFunctions.Invocation} invocation Invocation object.
 * @returns The value of the cell at the input address.
 **/ 
export async function getRangeExcelContextSet(address, invocation) {
  const context = new Excel.RequestContext();
  context.setInvocation(invocation);

  const range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
  range.load("values");

  await context.sync(); 
  return range.values[0][0];
}
```

## See also

- [Autogenerate JSON metadata for custom functions](custom-functions-json-autogeneration.md)
- [Manually create JSON metadata for custom functions](custom-functions-json.md)
