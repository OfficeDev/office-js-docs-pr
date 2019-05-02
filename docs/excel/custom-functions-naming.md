---
ms.date: 05/02/2019
description: Learn requirements for Excel custom functions' names and avoid common naming pitfalls.
title: Naming guidelines for custom functions in Excel
localization_priority: Normal
---
# Naming guidelines

A custom function is identified by an **id** and **name** property in the JSON metadata file. 

- The function `id` is used to uniquely identify custom functions in your JavaScript code. 
- The function `name` is used as the display name that appears to a user in Excel. 

A function `name` can differ from the function `id`, such as for localization purposes. In general, a function's `name` should stay the same as the `id` if there is no compelling reason for them to differ.

A function's `name` and `id` share some common requirements:

- A function's `id` may only use characters A through Z, numbers zero through nine, underscores, and periods.

- A function's `name` may use any Unicode alphabetic characters, underscores, and periods.

- Both function `name` and `id` must start with a letter and have a minimum limit of three characters.

Excel uses uppercase letters for built-in function names (such as `SUM`). Therefore, consider using uppercase letters for your custom function's `name` and `id` as a best practice.

A function's `name` shouldn't be named the same as:

- Any cells between A1 to XFD1048576 or any cells between R1C1 to R1048576C16384.

- Any Excel 4.0 Macro Function (such as `RUN`, `ECHO`).  For a full list of these functions, see [this article](https://www.microsoft.com/en-us/download/details.aspx?id=1465).

## Naming conflicts

If your function `name` is the same as a function `name` in an add-in that already exists, the **#REF!** error will appear in your workbook.

To fix a naming conflict, change the `name` in your add-in and try the function again. You can also uninstall the add-in with the conflicting name. Or, if you're testing your add-in in different environments, try using a different namespace to differentiate your function (such as `NAMESPACE_NAMEOFFUNCTION`).

## Best practices

- Consider adding multiple arguments to a function rather than creating multiple functions with the same or similar names.
- Function names should indicate the action of the function, such as `=GETZIPCODE` instead of `ZIPCODE`.
- Avoid ambiguous abbreviations in function names. Clarity is more important than brevity. Choose a name like `=INCREASETIME` rather than `=INC`.
- Consistently use the same verbs for functions which perform similar actions. For example, use `=DELETEZIPCODE` and `=DELETEADDRESS`, rather than `=DELETEZIPCODE` and `=REMOVEADDRESS`.

## Localizing function names

You can localize your function names for different languages using separate JSON files and override values in your add-in's manifest file. As a best practice, avoid giving your functions an `id` or `name` that is a built-in Excel function in another language as this could conflict with localized functions.

For full information on localizing, see [Localize custom functions](custom-functions-localize.md)

## Next steps
Learn about [error handling best practices](#custom-functions-errors.md).

## See also

* [Custom functions metadata](custom-functions-json.md)
* [Custom functions best practices](custom-functions-best-practices.md)
* [Excel custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md)
* [Runtime for Excel custom functions](custom-functions-runtime.md)
