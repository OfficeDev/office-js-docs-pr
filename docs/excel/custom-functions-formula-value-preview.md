---
title: Work with formula value preview mode in your custom functions
description: Work with formula value preview mode in your custom functions
ms.date: 05/15/2025
ms.localizationpriority: medium
---

# Work with formula value preview mode in your custom functions

You can control how your custom function calculates results when participating in formula value preview mode. Formula value preview mode is a feature that allows end users to select portions of a formula while editing the cell to preview the values. This feature helps users evaluate the formula as they edit it. The following image shows an example of the user editing a formula and selecting the text `A1+A2`. The formula preview mode shows the value `7` above.

:::image type="content" source="../images/excel-formula-value-preview.png" alt-text="Screenshot of Excel formula editor with A1+A2 selected and a preview value of 7 displayed above the formula editor.":::

By default, custom functions (for example `=getHousePrice(A1)`) can be previewed by the user. However, the following list shows some scenarios in which you may want to control how your custom function participates in formula value preview mode.

- Your custom function calls one or more APIs that charge a rate for using them.
- Your custom function accesses one or more scarce resources such as databases.
- Your custom function takes significant time to calculate the result, and it wouldn’t be useful for a user during preview purposes.

You can change the behavior of your custom function to return a mock value instead. To do this use the `invocation.isInValuePreview` read-only property. The following code sample shows an example custom function named `getHousePrice` that looks up house prices through a monetized API. If `isInValuePreview` is `true`, the custom function returns a mock number to be used and avoids incurring any cost. If `isInValuePreview` is `false`, the custom function calls the API and returns the actual house price value for use in the Excel spreadsheet.

```javascript
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
    // Return a useable but fake number.
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
