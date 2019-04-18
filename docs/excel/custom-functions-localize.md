---
ms.date: 03/15/2019
description: Localize your Excel custom functions. 
title: Localization of Excel custom functions (preview)
localization_priority: Normal
---
# Localize custom functions

To make your custom functions work around the world, localize them into different languages. To localize custom functions, you'll need to provide locale information in the XML manifest file and provide localized function names in the functions' JSON file.

## Localize your add-in

Before localizing your function names, you'll need to add optional locales to your XML file. For more information on this process, see [Localization for Office Add-ins](../develop/localization.md#control-localization-from-the-manifest).

## Localize function names

To declare a localized function name, create additional `name` and `description` properties within your function's JSON file. You can declare multiple localized names and descriptions.

The `name` and `description` appear in Excel and are localized. However, the `id` of each function is not localized. The `id` property is how Excel identifies your function as unique and should not be changed once it is set.

In the following code sample, you'll see the JSON file for a function with the `id` property "MULTIPLY." The `name` and `description` property of the function is localized for German. Each parameter `name` and `description` is also localized for German.

```JSON
{
    "id": "MULTIPLY",
    "name": "SUMME",
    "description": "Summe zwei Zahlen",
    "helpUrl": "http://www.contoso.com",
    "result": {
        "type": "number",
        "dimensionality": "scalar"
    },
    "parameters": [
        {
            "name": "eins",
            "description": "Erste Nummer",
            "dimensionality": "scalar"
        },
        {
            "name": "zwei",
            "description": "Zweite Nummer",
            "dimensionality": "scalar"
        },
    ],
}
```

## See also

* [Create custom functions in Excel](custom-functions-overview.md)
* [Custom functions metadata](custom-functions-json.md)
* [Custom functions best practices](custom-functions-best-practices.md)
* [Custom functions changelog](custom-functions-changelog.md)
* [Excel custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md)
