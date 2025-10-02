---
title: Delay execution while cell is being edited
description: Defer an Excel.run batch until the user leaves cell edit mode instead of failing with an error.
ms.date: 09/19/2025
ms.localizationpriority: medium
---

# Delay execution while cell is being edited

When a user is actively editing a cell, some add-in operations may fail immediately. Use the `delayForCellEdit` option to queue the batch until the user exits cell edit mode instead of throwing an error.

`Excel.run` has an overload that takes an [Excel.RunOptions](/javascript/api/excel/excel.runoptions) object. It supports the following property relevant to this scenario.

- `delayForCellEdit`: When `true`, Excel queues the batch until the user exits cell edit mode. When `false`, the batch fails immediately if the user is editing. The default value is `false`.

## Behavior comparison

| User editing? | `delayForCellEdit = false` (default) | `delayForCellEdit = true` |
|---------------|------------------------------------|--------------------------|
| No | Batch runs immediately | Batch runs immediately |
| Yes | Batch fails (`InvalidOperation` error) | Batch waits; then runs after edit is committed or canceled |

## Example

```js
await Excel.run({ delayForCellEdit: true }, async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("A1");
  range.values = [["Updated while user was editing elsewhere"]];
  await context.sync();
});
```

## Guidance

- Use `delayForCellEdit` only when user-initiated commands may overlap active cell editing, such as with a ribbon button that triggers a bulk update.
- Consider showing a status indicator in your add-in if queued work may be lengthy.
- Avoid chaining multiple long-running delayed batches. This creates a perceived lag for your users.

## Next steps

Explore related user-state strategies in [Events](excel-add-ins-events.md), such as reacting to selection changes, and combine with `delayForCellEdit` to improve your add-in's robustness.
