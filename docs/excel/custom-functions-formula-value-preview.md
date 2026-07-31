---
title: Work with formula value preview mode in your custom functions
description: Control whether an Excel custom function returns a preview value or runs its full calculation while a user edits a formula.
ms.date: 07/28/2026
ms.topic: how-to
ms.localizationpriority: medium
ai-usage: ai-assisted
---

# Work with formula value preview mode in your custom functions

Formula value preview mode helps users evaluate a formula while they edit it. When a user selects part of a formula, Excel calculates and displays the selected value. For example, the following image shows a preview value of `7` for the selected expression `A1+A2`.

:::image type="content" source="../images/excel-formula-value-preview.png" alt-text="Screenshot of Excel formula editor with A1+A2 selected and a preview value of 7 displayed above the formula editor.":::

By default, Excel runs a custom function during formula value preview. This behavior might cause problems depending on what your custom function does. Use the read-only `invocation.isInValuePreview` property to detect a preview calculation and return a mock value when the full calculation would:

- Call a metered API.
- Access a limited resource, such as a database.
- Take too long to provide a useful preview.

The following `getHousePrice` custom function returns a mock price during preview. For a standard calculation, it calls the metered service and returns the actual price.

```typescript
/**
 * Get the listing price for a house on the market for the given address.
 * @customfunction
 * @param address The address of the house.
 * @param invocation Custom function handler.
 * @returns The price of the house at the address.
 */
export function getHousePrice(address: string, invocation: CustomFunctions.Invocation): number {
  // Check if this call is for formula value preview mode.
  if (invocation.isInValuePreview) {
    // Avoid long-running expensive service calls.
    // Return a usable but fake number.
    return 450000;
  } else {
    // Make the actual service calls in this block.
    const price = callHouseServiceAPI(address);
    return price;
  }
}
```

## See also

- [Create custom functions in Excel](custom-functions-overview.md)
- [Custom functions parameter options](custom-functions-parameter-options.md)
