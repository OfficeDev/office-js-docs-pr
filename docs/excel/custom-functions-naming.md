---
ms.date: 02/08/2019
description: Learn requirements for Excel custom functions' names and avoid common naming pitfalls.
title: Naming guidelines for custom functions in Excel (preview)
localization_priority: Normal
---
# Naming guidelines

Custom functions contain two unique properties in their JSON metadata file: their function name and their function ID. Function IDs are used as the main way to differentiate functions from each other in Excel programmatically. Function names are what appear to users of your add-in as the function's name in Excel. Function names can differ from the ID, e.g. for localization purposes, but should generally stay the same as the ID if there is not a compelling reason for them to differ.

Function names and function IDs share some common requirements:

- Can use alphanumeric characters (including Unicode), the numbers zero through nine, underscores, and periods.

- Must start with a letter and have a minimum limit of three characters.

Most function names for built-in Excel functions appear in uppercase letters (e.g. `SUM`). While not required, it is a best practice to specify your custom function names and function IDs using uppercase letters.

Function names should not be named the same as:

- Any cells between A1 to XFD1048576 or any cells between R1C1 to R1048576C16384.

- Any Excel 4.0 Macro Function (such as `RUN`, `ECHO`).  For a full list of these functions, see [this article](https://www.microsoft.com/en-us/download/details.aspx?id=1465).

## Naming conflicts

If a function name in your add-in matches a name in an existing and installed add-in, you'll see an error #REF! appear in your workbook.

To troubleshoot this, please change the name in your add-in and try the function again. If you're testing your add-in in different environments, either uninstall the add-in with the conflicting name or use a different namespace to differentiate your function (e.g. NAMESPACE_NAMEOFFUNCTION).

Also consider how you'd like people to use the functions within your add-in. In many cases, it makes sense to add multiple arguments to a function rather than create multiple functions with the same or similar names.