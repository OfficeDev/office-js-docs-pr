---
ms.date: 01/14/2020
description: Define JSON metadata for custom functions in Excel and associate your function id and name properties.
title: Metadata for custom functions in Excel
localization_priority: Normal
---

# Custom functions metadata

As described in the [custom functions overview](custom-functions-overview.md) article, a custom functions project must include both a JSON metadata file and a script (either JavaScript or TypeScript) file to register a function, making it available for use. Custom functions are registered when the user runs the add-in for the first time and after that are available to the same user in all workbooks.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

It is recommended that you use JSON autogeneration when possible, using the `yo office` scaffold files, similar to the process shown in the [Excel Custom Function tutorial](../tutorials/excel-tutorial-create-custom-functions.md) because this process is easier and less prone to user error. For more information on the process of JSDoc comment JSON file generation, see [Generate JSON metadata for custom functions](custom-functions-json-autogeneration.md).

However, you can make a custom functions project from scratch; it requires that you:

- Write your JSON file by hand
- Check that your manifest file is connected to your hand-authored JSON file
- Associate your functions' `id` and `name` properties in the script file in order to register your functions

This article will show you how to do all three of these steps.

The following image explains the differences between using `yo office` scaffold files and writing JSON from scratch.
![Image of differences between using Yo Office and writing your own JSON](../images/custom-functions-json.png)

> [!NOTE]
> In contrast with the `yo office` scaffold files, you need to connect your manifest to the JSON file you create, through the `<Resources>` section in your XML manifest file. Note that the server settings on the server that hosts the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel on the web.

## Authoring metadata and connecting to the manifest

You need to create a JSON file in your project and provide all the details about your functions in it, such as the function's parameters. See the [following metadata example](#json-metadata-example) and [the metadata reference](#metadata-reference) for a complete list of function properties.

You also need to make sure your XML manifest file references your JSON file in the `<Resources>` section, similar to the following example.

```json
<Resources>
    <bt:Urls>
        <bt:Url id="JSON-URL" DefaultValue="https://subdomain.contoso.com/config/customfunctions.json"/>
        <bt:Url id="JS-URL" DefaultValue="https://subdomain.contoso.com/dist/win32/ship/index.win32.bundle"/>
            <bt:Url id="HTML-URL" DefaultValue="https://subdomain.contoso.com/index.html"/>
    </bt:Urls>
    <bt:ShortStrings>
        <bt:String id="namespace" DefaultValue="CONTOSO"/>
    </bt:ShortStrings>
</Resources>
```

## JSON metadata example

The following example shows the contents of a JSON metadata file for an add-in that defines custom functions. The sections that follow this example provide detailed information about the individual properties within this JSON example.

```json
{
  "functions": [
    {
      "id": "ADD",
      "name": "ADD",
      "description": "Add two numbers",
      "helpUrl": "http://www.contoso.com/help",
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
          "dimensionality": "scalar"
        }
      ]
    },
    {
      "id": "GETDAY",
      "name": "GETDAY",
      "description": "Get the day of the week",
      "helpUrl": "http://www.contoso.com/help",
      "result": {
        "dimensionality": "scalar"
      },
      "parameters": []
    },
    {
      "id": "INCREMENTVALUE",
      "name": "INCREMENTVALUE",
      "description": "Count up from zero",
      "helpUrl": "http://www.contoso.com/help",
      "result": {
        "dimensionality": "scalar"
      },
      "parameters": [
        {
          "name": "increment",
          "description": "the number to be added each time",
          "type": "number",
          "dimensionality": "scalar"
        }
      ],
      "options": {
        "stream": true,
        "cancelable": true
      }
    },
    {
      "id": "SECONDHIGHEST",
      "name": "SECONDHIGHEST",
      "description": "Get the second highest number from a range",
      "helpUrl": "http://www.contoso.com/help",
      "result": {
        "dimensionality": "scalar"
      },
      "parameters": [
        {
          "name": "range",
          "description": "the input range",
          "type": "number",
          "dimensionality": "matrix"
        }
      ]
    }
  ]
}
```

> [!NOTE]
> A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json) GitHub repository's commit history. As the project has been adjusted to automatically generate JSON, a full sample of handwritten JSON is only available in previous versions of the project.

## Metadata reference

### functions

The `functions` property is an array of custom function objects. The following table lists the properties of each object.

| Property      | Data type | Required | Description                                                                                                                                                                      |
| :------------ | :-------- | :------- | :------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `description` | string    | No       | The description of the function that end users see in Excel. For example, **Converts a Celsius value to Fahrenheit**.                                                            |
| `helpUrl`     | string    | No       | URL that provides information about the function. (It is displayed in a task pane.) For example, `http://contoso.com/help/convertcelsiustofahrenheit.html`.                      |
| `id`          | string    | Yes      | A unique ID for the function. This ID can only contain alphanumeric characters and periods and should not be changed after it is set.                                            |
| `name`        | string    | Yes      | The name of the function that end users see in Excel. In Excel, this function name will be prefixed by the custom functions namespace that's specified in the XML manifest file. |
| `options`     | object    | No       | Enables you to customize some aspects of how and when Excel executes the function. See [options](#options) for details.                                                          |
| `parameters`  | array     | Yes      | Array that defines the input parameters for the function. See [parameters](#parameters) for details.                                                                             |
| `result`      | object    | Yes      | Object that defines the type of information that is returned by the function. See [result](#result) for details.                                                                 |

### options

The `options` object enables you to customize some aspects of how and when Excel executes the function. The following table lists the properties of the `options` object.

| Property          | Data type | Required                               | Description                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                |
| :---------------- | :-------- | :------------------------------------- | :--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `cancelable`      | boolean   | No<br/><br/>Default value is `false`.  | If `true`, Excel calls the `CancelableInvocation` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function. Cancelable functions are typically only used for asynchronous functions that return a single result and need to handle the cancellation of a request for data. A function cannot be both streaming and cancelable. For more information, see the note near the end of [Make a streaming function](custom-functions-web-reqs.md#make-a-streaming-function). |
| `requiresAddress` | boolean   | No <br/><br/>Default value is `false`. | If `true`, your custom function can access the address of the cell that invoked your custom function. To get the address of the cell that invoked your custom function, use context.address in your custom function. For more information, see [Addressing cell's context parameter](/office/dev/add-ins/excel/custom-functions-parameter-options#addressing-cells-context-parameter). Custom functions cannot be set as both streaming and requiresAddress. When using this option, the 'invocation' parameter must be the last parameter passed in options.                                              |
| `stream`          | boolean   | No<br/><br/>Default value is `false`.  | If `true`, the function can output repeatedly to the cell even when invoked only once. This option is useful for rapidly-changing data sources, such as a stock price. The function should have no `return` statement. Instead, the result value is passed as the argument of the `StreamingInvocation.setResult` callback method. For more information, see [Streaming functions](custom-functions-web-reqs.md#make-a-streaming-function).                                                                                                                                                                |
| `volatile`        | boolean   | No <br/><br/>Default value is `false`. | <br /><br /> If `true`, the function will recalculate each time Excel recalculates, instead of only when the formula's dependent values have changed. A function cannot be both streaming and volatile. If the `stream` and `volatile` properties are both set to `true`, the volatile option will be ignored.                                                                                                                                                                                                                                                                                             |

### parameters

The `parameters` property is an array of parameter objects. The following table lists the properties of each object.

|  Property  |  Data type  |  Required  |  Description  |
|:-----|:-----|:-----|:-----|
|  `description`  |  string  |  No |  A description of the parameter. This is displayed in Excel's intelliSense.  |
|  `dimensionality`  |  string  |  No  |  Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).  |
|  `name`  |  string  |  Yes  |  The name of the parameter. This name is displayed in Excel's intelliSense.  |
|  `type`  |  string  |  No  |  The data type of the parameter. Can be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types. If this property is not specified, the data type defaults to **any**. |
|  `optional`  | boolean | No | If `true`, the parameter is optional. |
|`repeating`| boolean | No | If `true`, parameters will populate from a specified array. Note that functions all repeating parameters are considered optional parameters by definition.  |

### result

The `result` object defines the type of information that is returned by the function. The following table lists the properties of the `result` object.

| Property         | Data type | Required | Description                                                                          |
| :--------------- | :-------- | :------- | :----------------------------------------------------------------------------------- |
| `dimensionality` | string    | No       | Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array). |

## Associating function names with JSON metadata

For a function to work properly, you need to associate the function's `id` property with the JavaScript implementation. Make sure there is an association, otherwise the function will not be registered and not useable in Excel. The following code sample shows how to make the association using the `CustomFunctions.associate()` method. The sample defines the custom function `add` and associates it with the object in the JSON metadata file where the value of the `id` property is **ADD**.

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
    }
  ]
}
```

Keep in mind the following best practices when creating custom functions in your JavaScript file and specifying corresponding information in the JSON metadata file.

- In the JSON metadata file, ensure that the value of each `id` property contains only alphanumeric characters and periods.

- In the JSON metadata file, ensure that the value of each `id` property is unique within the scope of the file. That is, no two function objects in the metadata file should have the same `id` value.

- Do not change the value of an `id` property in the JSON metadata file after it's been associated with a corresponding JavaScript function name. You can change the function name that end users see in Excel by updating the `name` property within the JSON metadata file, but you should never change the value of an `id` property after it's been established.

- In the JavaScript file, specify a custom function association using `CustomFunctions.associate` after each function.

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

## Next steps

Learn the [best practices for naming your function](custom-functions-naming.md) or discover how to [localize your function](custom-functions-localize.md) using the previously described handwritten JSON method.

## See also

- [Autogenerate JSON metadata for custom functions](custom-functions-json-autogeneration.md)
- [Custom functions parameter options](custom-functions-parameter-options.md)
- [Create custom functions in Excel](custom-functions-overview.md)
