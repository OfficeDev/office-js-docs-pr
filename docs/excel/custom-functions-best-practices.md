---
ms.date: 01/08/2019
description: Learn best practices for developing custom functions in Excel.
title: Custom functions best practices (preview)
localization_priority: Normal
---

# Custom functions best practices (preview)

This article describes best practices for developing custom functions in Excel.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## Troubleshooting

1. If you are testing your add-in in Office on Windows, you should enable **[runtime logging](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in)** to troubleshoot issues with your add-in's XML manifest file, as well as several installation and runtime conditions. Runtime logging writes `console.log` statements to a log file to help you uncover issues.

2. Your add-in will not load if one or more custom functions conflicts with a previously registered add-in's custom functions. In this case, you can either remove the existing add-in, or if you encounter this error while developing an add-in, you can specify a different namespace name in your manifest.

3. To report feedback to the Excel Custom Functions team about this method of troubleshooting, send the team feedback. To do this, select **File | Feedback | Send a Frown**. Sending a frown will provide the necessary logs to understand the issue you are hitting.

## Debugging

Currently, the best method for debugging Excel custom functions is to first [sideload](../testing/sideload-office-add-ins-for-testing.md) your add-in within **Excel Online**. You can then debug your custom functions by using the [F12 debugging tool native to your browser](../testing/debug-add-ins-in-office-online.md) in combination with the following techniques:

- Use `console.log` statements within your custom functions code to send output to the console in real time.

- Use `debugger;` statements within your custom functions code to specify breakpoints where execution will pause when the F12 window is open. For example, if the following function runs while the F12 window is open, execution will pause on the `debugger;` statement, enabling you to manually inspect parameter values before the function returns. The `debugger;` statement has no effect in Excel Online when the F12 window is not open. Currently, the `debugger;` statement has no effect in Excel for Windows.

    ```js
    function add(first, second){
      debugger;
      return first + second;
    }
    ```

If your add-in fails to register, [verify that SSL certificates are correctly configured](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) for the web server that's hosting your add-in application.

## Associating function names with JSON metadata

As described in the [custom functions overview](custom-functions-overview.md) article, a custom functions project must include both a JSON metadata file and a script (either JavaScript or TypeScript) file to form a complete function. For a function to work properly, you'll need to bind the name of the function in the script file to the id listed in the JSON file. This process is called association. Make a note to include associations at the end of your JavaScript code files; otherwise, your functions will not work.

The following code sample shows how to do this association. The sample defines the custom function `add` and associates it with the object in the JSON metadata file where the value of the `id` property is **ADD**.

```js
function add(first, second){
  return first + second;
}

CustomFunctions.associate("ADD", add);
```

Keep in mind the following best practices when creating custom functions in your JavaScript file and specifying corresponding information in the JSON metadata file.

* Only use uppercase letters for a function's `name` and `id` in the JSON metadata file. Do not use a mix of cases or only lowercase letters. If you do, you may end up with two values that only differ by case which will cause unintentional overwriting of your functions. For example, a function object with an `id` value of **add** could be overwritten by declaration later in the file of function object with an `id` value of **ADD**. Additionally, the `name` property defines the function name that end users will see in Excel. Using uppercase letters for the name of each custom function provides a consistent experience in Excel, where all built-in function names are uppercase.

* However, it is not necessary to capitalize the function's `name` when associating. For example, `CustomFunctions.associate("add", add)` is equivalent to `CustomFunctions.associate("ADD", add)`.

* In the JSON metadata file, ensure that the value of each `id` property contains only alphanumeric characters and periods.

* In the JSON metadata file, ensure that the value of each `id` property is unique within the scope of the file. That is, no two function objects in the metadata file should have the same `id` value. 

* Do not change the value of an `id` property in the JSON metadata file after it's been associated with a corresponding JavaScript function name. You can change the function name that end users see in Excel by updating the `name` property within the JSON metadata file, but you should never change the value of an `id` property after it's been established.

* In the JavaScript file, specify all custom function associations in the same location. For example, the following code sample defines two custom functions and then specifies the association information for both functions.

    ```js
    function add(first, second){
      return first + second;
    }

    function increment(incrementBy, callback) {
      var result = 0;
      var timer = setInterval(function() {
        result += incrementBy;
        callback.setResult(result);
      }, 1000);

      callback.onCanceled = function() {
        clearInterval(timer);
      };
    }

    // associate `id` values in the JSON metadata file to JavaScript function names
    CustomFunctions.associate("ADD", add);
    CustomFunctions.associate("INCREMENT", increment);
    ```

    The following sample shows the JSON metadata that corresponds to the functions defined in this JavaScript code sample. Note that the `id` and `name` properties are in uppercase letters in this file. 

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

## Declaring optional parameters 

In Excel for Windows (version 1812 or later), you can declare optional parameters for your custom functions. When a user invokes a function in Excel, optional parameters appear in brackets. For example, a function `FOO` with one required parameter called `parameter1` and one optional parameter called `parameter2` would appear as `=FOO(parameter1, [parameter2])` in Excel.

To make a parameter optional, add `"optional": true` to the parameter in the JSON metadata file that defines the function. The following example shows what this might look like for the function `=ADD(first, second, [third])`. Notice that the optional `[third]` parameter follows the two required parameters. Required parameters will appear first in Excel’s Formula UI.

```json
{
    "id": "ADD",
    "name": "ADD",
    "description": "Add two numbers",
    "helpUrl": "http://www.contoso.com",
    "result": {
        "type": "number",
        "dimensionality": "scalar"
        },
    "parameters": [
        {
            "name": "first",
            "description": "first number to add",
            "type": "number",
            "dimensionality": "scalar"
        },
        {
            "name": "second",
            "description": "second number to add",
            "type": "number",
            "dimensionality": "scalar",
        },
        {
            "name": "third",
            "description": "third optional number to add",
            "type": "number",
            "dimensionality": "scalar",
            "optional": true
        }
    ],
    "options": {
        "sync": false
    }
}
```

When you define a function that contains one or more optional parameters, you should specify what happens when the optional parameters are undefined. In the following example, `zipCode` and `dayOfWeek` are both optional parameters for the `getWeatherReport` function. If the `zipCode` parameter is undefined, the default value is set to 98052. If the `dayOfWeek` parameter is undefined, it is set to Wednesday.

```js
function getWeatherReport(zipCode, dayOfWeek)
{
  if (zipCode === undefined) {
      zipCode = "98052";
  }

  if (dayOfWeek === undefined) {
    dayOfWeek = "Wednesday";
  }

  // Get weather report for specified zipCode and dayOfWeek
  // ...
}
```

## Additional considerations

In order to create an add-in that will run on multiple platforms (one of the key tenants of Office Add-ins), you should not access the Document Object Model (DOM) in custom functions or use libraries like jQuery that rely on the DOM. On Excel for Windows, where custom functions use the [JavaScript runtime](custom-functions-runtime.md), custom functions cannot access the DOM.

## See also

* [Create custom functions in Excel](custom-functions-overview.md)
* [Custom functions metadata](custom-functions-json.md)
* [Runtime for Excel custom functions](custom-functions-runtime.md)
* [Custom functions changelog](custom-functions-changelog.md)
* [Excel custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md)
