---
title: Naming guidelines for custom functions in Excel
description: Learn requirements for names of Excel custom functions and avoid common naming pitfalls.
ms.date: 07/08/2021
ms.localizationpriority: medium
---
# Custom functions naming guidelines

A custom function is identified by an `id` and `name` property in the JSON metadata file.

- The function `id` is used to uniquely identify custom functions in your JavaScript code.
- The function `name` is used as the display name that appears to a user in Excel.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

A function `name` can differ from the function `id`, such as for localization purposes. In general, a function's `name` should stay the same as the `id` if there is no reason for them to differ.

A function's `name` and `id` share some common requirements.

- A function's `id` may only use characters A through Z, numbers zero through nine, underscores, and periods.

- A function's `name` may use any Unicode alphabetic characters, underscores, and periods.

- Both function `name` and `id` must start with a letter and have a minimum limit of three characters.

Excel uses uppercase letters for built-in function names (such as `SUM`). Use uppercase letters for your custom function's `name` and `id` as a best practice.

A function's `name` shouldn't be the same as:

- Any cells between A1 to XFD1048576 or any cells between R1C1 to R1048576C16384.

- Any Excel 4.0 Macro Function (such as `RUN`, `ECHO`).  For a full list of these functions, see [this Excel Macro Functions Reference document](https://www.myonlinetraininghub.com/cdn/files/Excel%204.0%20Macro%20Functions%20Reference.pdf).

## Naming conflicts

If your function `name` is the same as a function `name` in an add-in that already exists, the **#REF!** error will appear in your workbook.

To fix a naming conflict, change the `name` in your add-in and try the function again. You can also uninstall the add-in with the conflicting name. Or, if you're testing your add-in in different environments, try using a different namespace to differentiate your function (such as `NAMESPACE_NAMEOFFUNCTION`).

## Best practices

- Consider adding multiple arguments to a function rather than creating multiple functions with the same or similar names.
- Avoid ambiguous abbreviations in function names. Clarity is more important than brevity. Choose a name like `=INCREASETIME` rather than `=INC`.
- Function names should indicate the action of the function, such as =GETZIPCODE instead of ZIPCODE.
- Consistently use the same verbs for functions which perform similar actions. For example, use `=DELETEZIPCODE` and `=DELETEADDRESS`, rather than `=DELETEZIPCODE` and `=REMOVEADDRESS`.
- When naming a streaming function, consider adding a note to that effect in the description of the function or adding `STREAM` to the end of the function's name.

[!include[manifest guidance](../includes/manifest-guidance.md)]

## Localizing function names

You can localize your function names for different languages using separate JSON files and override values in your add-in's manifest file. Avoid giving your functions an `id` or `name` that is a built-in Excel function in another language as this could conflict with localized functions.

For full information on localizing, see [Localize custom functions](custom-functions-localize.md)

## Next steps

Learn about [error handling best practices](custom-functions-errors.md).

## See also

* [Manually create JSON metadata for custom functions](custom-functions-json.md)
* [Excel custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md)
