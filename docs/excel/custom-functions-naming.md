---
ms.date: 02/08/2019
description: Learn requirements for Excel custom functions' names and avoid common naming pitfalls.
title: Naming guidelines for custom functions in Excel (preview)
localization_priority: Normal
---
# Naming guidelines

A custom function is identified by an **id** and **name** property in the JSON metadata file. The function id is used to uniquely identify custom functions in your JavaScript code. The function name is used as the display name that appears to a user in Excel. A function name can differ from the function ID, such as for localization purposes. But in general it should stay the same as the ID if there is no compelling reason for them to differ.

Function names and function IDs share some common requirements:

- Function ids may only use characters A through Z, numbers zero through nine, underscores, and periods.

- Function names may use any Unicode alphabetic characters, underscores, and periods.

- They must start with a letter and have a minimum limit of three characters.

Excel uses uppercase letters for built-in function names (such as `SUM`). Therefore, consider using uppercase letters for your custom function names and function IDs as a best practice.

Function names shouldn't be named the same as:

- Any cells between A1 to XFD1048576 or any cells between R1C1 to R1048576C16384.

- Any Excel 4.0 Macro Function (such as `RUN`, `ECHO`).  For a full list of these functions, see [this article](https://www.microsoft.com/en-us/download/details.aspx?id=1465).

## Naming conflicts

If your function name is the same as a function name in an add-in that already exists, the **#REF!** error will appear in your workbook.

To fix a name conflict, change the name in your add-in and try the function again. You can also uninstall the add-in with the conflicting name. Or, if you're testing your add-in in different environments, try using a different namespace to differentiate your function (such as NAMESPACE_NAMEOFFUNCTION).

Also consider how you'd like people to use the functions within your add-in. In many cases, it makes sense to add multiple arguments to a function rather than create multiple functions with the same or similar names.

## See also

* [Custom functions metadata](custom-functions-json.md)
* [Custom functions best practices](custom-functions-best-practices.md)
* [Excel custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md)
* [Runtime for Excel custom functions](custom-functions-runtime.md)
