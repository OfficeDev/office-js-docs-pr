---
title: Manage custom function visibility
description: Show or hide custom functions from the Excel UI.
ms.date: 08/10/2026
ms.topic: how-to
ms.localizationpriority: medium
---

# Show or hide custom functions in the Excel UI

Control which custom functions display in Excel AutoComplete and the Formula Builder. If your add-in serves multiple user types (such as parents, teachers, and students), each category can use a specialized set of custom functions.

> [!NOTE]
> To hide custom functions before an add-in launches, use the [`excludeFromAutoComplete` JSDoc tag](custom-functions-json-autogeneration.md#excludeFromAutoComplete) or set the [`excludeFromAutoComplete` property](custom-functions-json.md#options) to `true`.

The following code sample maps functions to different categories of add-in users so that the functions are programmatically visible or hidden for each user type. `FUNCTIONBASIC`, `FUNCTIONA`, `FUNCTIONB`, and `FUNCTIONC` are the exact short names from the functions' JSON metadata.

```typescript
/**
 * This code snippet maps existing custom functions to add-in user types.
 * The primary function, FUNCTIONBASIC, is visible for all user types.
 * The other three functions are only visible to specific user types.
 */
const allFunctions = [
    "FUNCTIONBASIC",
    "FUNCTIONA",
    "FUNCTIONB",
    "FUNCTIONC",
];

// Assign each function to a user type.
const userFunctionMapping = new Map<string, string[]>([
    ["banker", ["FUNCTIONBASIC", "FUNCTIONA", "FUNCTIONB"]],
    ["trader", ["FUNCTIONBASIC", "FUNCTIONB"]],
    ["analyst", ["FUNCTIONBASIC", "FUNCTIONA", "FUNCTIONC"]],
]);

// Create a placeholder to retrieve the current user type.
(async () => {
    await Office.onReady();
    const userType = getCurrentUser(); // Implement `getCurrentUser()` to return the current user type (banker, trader, or analyst).
    await showFunctionsBasedOnUserType(userType);
})();

// Show the correct functions based on the current user type.
async function showFunctionsBasedOnUserType(userType: string) {
    const functionsToShow = userFunctionMapping.get(userType) ?? [];
    const functionsToHide = allFunctions.filter((name) => !functionsToShow.includes(name));
    const customFunctionVisibilityOptions: Excel.CustomFunctionVisibilityOptions = {
        show: functionsToShow,
        hide: functionsToHide,
    };
    await Excel.CustomFunctionManager.setVisibility(customFunctionVisibilityOptions);
}
```

> [!IMPORTANT]
> Values in the `show` and `hide` arrays must exactly match the short custom function names in the generated or manually authored JSON metadata. Matching is case-sensitive, so preserve the capitalization in the metadata. Don't include the manifest namespace. For example, if Excel displays `CONTOSO.CLOCK` and the metadata name is `CLOCK`, pass `"CLOCK"`, not `"CONTOSO.CLOCK"`. Don't assume the exported JavaScript or TypeScript implementation name is the required value. For example, `logMessage` in the preceding metadata has the short name `LOG`, so pass `"LOG"`.

The `show` array reveals the listed functions, and the `hide` array hides the listed functions. Omitting a function from `show` doesn't automatically hide it. Visibility affects AutoComplete and the Formula Builder only. It doesn't unregister a custom function or prevent a user from entering its formula manually.

## See also

- [Manually create JSON metadata for custom functions](custom-functions-json.md)
- [Autogenerate JSON metadata for custom functions](custom-functions-json-autogeneration.md)
- [What's new in Excel JavaScript API 1.20](/javascript/api/requirement-sets/excel/excel-api-1-20-requirement-set)
