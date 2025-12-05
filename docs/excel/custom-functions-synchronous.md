---
title: Synchronous custom functions
description: Show or hide custom functions from the Excel UI.
ms.date: 12/05/2025
ms.localizationpriority: medium
---

# Synchronous custom functions

Synchronous custom functions support evaluate and conditional format actions in Excel. To support synchronous scenarios, a custom function must use the @supportSync JSDoc tag or supportSync: true setting in the functions.json file. Without an explicit synchronous setting, custom functions can't trigger evaluate or conditional format actions.

The following actions are supported with the synchronous settings.

## Evaluate

- UI action: Formulas > Evaluate Formula.
- UI action: Formulas > Insert Function.
- UI action: In cell edit mode, selecting part of a formula and using F9 to see partial calculation results.
- VBA API: Application.Calculate

## Conditional format

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