---
title: Manage custom function visibility
description: Show or hide custom functions from the Excel UI.
ms.date: 11/25/2025
ms.localizationpriority: medium
---

# Show or hide custom functions in the Excel UI

Control which custom functions display in Excel AutoComplete and the Formula Builder. If your add-in serves multiple user types (such as parents, teachers, and students), each category can use a specialized set of custom functions.

> [!NOTE]
> To hide custom functions before an add-in launches, use the [`excludeFromAutoComplete` JSDoc tag](custom-functions-json-autogeneration.md#excludeFromAutoComplete) or set the [`excludeFromAutoComplete` property](custom-functions-json.md#options) to `true`.

The following code sample shows how to map functions to different categories of add-in users so that the functions are programmatically visible or hidden for each user type. The sample assumes that four functions already exist, `functionBasic`, `functionA`, `functionB`, and `functionC`, and maps these functions to **banker**, **trader**, and **analyst** user types in an investment banking organization.

```typescript
/**
 * This code snippet maps existing custom functions to add-in user types.
 * The primary function, functionBasic, is visible for all user types. 
 * The other three functions, functionA, functionB, and functionC, are only visible to specific user types.
 */
const allFunctions = ["functionBasic", "functionA", "functionB", "functionC"];

// Assign each function to a user type.
const userFunctionMapping = new Map<string, string[]>([
    ["banker", ["functionBasic", "functionA", "functionB"]],
    ["trader", ["functionBasic", "functionB"]],
    ["analyst", ["functionBasic", "functionA", "functionC"]]
]);

// Create a placeholder to retrieve the current user type.
(async () => {
    await Office.onReady();
    let userType = getCurrentUser(); // Implement `getCurrentUser()` to return the current user type (banker, trader, or analyst).
    await showFunctionsBasedOnUserType(userType);
});

// Show the correct functions based on the current user type.
async function showFunctionsBasedOnUserType(userType: string) {
    let availableFunctions: string[] = userFunctionMapping.get(userType);
    let customFunctionVisibilityOptions: Excel.CustomFunctionVisibilityOptions = {
        show: availableFunctions,
    };
    // If functionA, functionB, and functionC are initially hidden with `excludeFromAutoComplete`, adjust visibility so the correct functions are shown.
    await Excel.CustomFunctionManager.setVisibility(customFunctionVisibilityOptions);
}
```

## See also

- [Manually create JSON metadata for custom functions](custom-functions-json.md)
- [Autogenerate JSON metadata for custom functions](custom-functions-json-autogeneration.md)
