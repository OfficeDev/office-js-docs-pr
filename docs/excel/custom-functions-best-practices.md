---
ms.date: 05/03/2019
description: Learn best practices for developing custom functions in Excel.
title: Custom functions best practices (preview)
localization_priority: Normal
---

# Custom functions best practices (preview)

This article describes best practices for developing custom functions in Excel.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## Associating function names with JSON metadata

As described in the [custom functions overview](custom-functions-overview.md) article, a custom functions project must include both a JSON metadata file and a script (either JavaScript or TypeScript) file to form a complete function. If you are using `yo office` the JSON metadata can be generated from the code comments. Otherwise you need to build the JSON metadata file manually.

For a function to work properly, you need to associate the id with the JavaScript implementation. Make sure there is an association, otherwise the function will not be called. The following code sample shows how to make the association using the `CustomFunctions.associate()` method. The sample defines the custom function `add` and associates it with the object in the JSON metadata file where the value of the `id` property is **ADD**.

```js
/**
 * Add two numbers
 * @customfunction
 * @param {number} first First number
 * @param {number} second Second number
 * @returns {number} The sum of the two numbers.
 */
function add(first, second) {
  return first + second;
}

CustomFunctions.associate("ADD", add);
```

The following JSON shows the JSON metadata that is associated with the previous custom function JavaScript code.

```json
{
  "functions": [
    {
        "description": "Add two numbers",
        "id": "ADD",
        "name": "ADD",
        "parameters": [
            {
                "description": "First number",
                "name": "first",
                "type": "number"
            },
            {
                "description": "Second number",
                "name": "second",
                "type": "number"
            }
        ],
        "result": {
            "type": "number"
        }
    },
  ]
}
```


Keep in mind the following best practices when creating custom functions in your JavaScript file and specifying corresponding information in the JSON metadata file.

* In the JSON metadata file, ensure that the value of each `id` property contains only alphanumeric characters and periods.

* In the JSON metadata file, ensure that the value of each `id` property is unique within the scope of the file. That is, no two function objects in the metadata file should have the same `id` value.

* Do not change the value of an `id` property in the JSON metadata file after it's been associated with a corresponding JavaScript function name. You can change the function name that end users see in Excel by updating the `name` property within the JSON metadata file, but you should never change the value of an `id` property after it's been established.

* In the JavaScript file, specify a custom function association using `CustomFunctions.associate` after each function.

The following sample shows the JSON metadata that corresponds to the functions defined in this JavaScript code sample. The `id` and `name` property values are in uppercase, which is a best practice when describing your custom functions. You only need to add this JSON if you are preparing your own JSON file manually and not using autogeneration. For more information on autogeneration, see [Create JSON metadata for custom functions](custom-functions-json-autogeneration.md).

```json
{
  "$schema": "https://developer.microsoft.com/en-us/json-schemas/office-js/custom-functions.schema.json",
  "functions": [
    {
      "id": "ADD",
      "name": "ADD",
      ...
    },
    {
      "id": "INCREMENT",
      "name": "INCREMENT",
      ...
    }
  ]
}
```

## Additional considerations

Avoid accessing the Document Object Model (DOM) directly or indirectly (for example, using jQuery) from your custom function. On Excel for Windows, where custom functions use the [JavaScript runtime](custom-functions-runtime.md), custom functions cannot access the DOM.

## Next steps
Learn how to [perform web requests with custom functions](custom-functions-web-reqs.md).

## See also

* [Autogenerate JSON metadata for custom functions](custom-functions-json-autogeneration.md)
* [Custom functions metadata](custom-functions-json.md)
* [Custom functions changelog](custom-functions-changelog.md)
* [Create custom functions in Excel](custom-functions-overview.md)