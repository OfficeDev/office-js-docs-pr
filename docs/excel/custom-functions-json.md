---
title: Manually create JSON metadata for custom functions in Excel
description: Define JSON metadata for custom functions in Excel and associate your function ID and name properties.
ms.date: 07/10/2025
ms.localizationpriority: medium
---

# Manually create JSON metadata for custom functions

As described in the [custom functions overview](custom-functions-overview.md) article, a custom functions project must include both a JSON metadata file and a script (either JavaScript or TypeScript) file to register a function, making it available for use. Custom functions are registered when the user runs the add-in for the first time and after that are available to the same user in all workbooks.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

We recommend using JSON autogeneration when possible instead of creating your own JSON file. Autogeneration is less prone to user error and the `yo office` scaffolded files already include this. For more information on JSDoc tags and the JSON autogeneration process, see [Autogenerate JSON metadata for custom functions](custom-functions-json-autogeneration.md).

However, you can make a custom functions project from scratch. This process requires you to:

- Write your JSON file.
- Check that your manifest file is connected to your JSON file.
- Associate your functions' `id` and `name` properties in the script file in order to register your functions.

The following image explains the differences between using `yo office` scaffold files and writing JSON from scratch.

![Image of differences between using the Yeoman generator for Office Add-ins and writing your own JSON.](../images/custom-functions-json.png)

> [!NOTE]
> Remember to connect your manifest to the JSON file you create, through the **\<Resources\>** section in your add-in only manifest file if you do not use the [Yeoman generator for Office Add-ins](../develop/yeoman-generator-overview.md).

## Authoring metadata and connecting to the manifest

Create a JSON file in your project and provide all the details about your functions in it, such as the function's parameters. See the [following metadata example](#json-metadata-example) and [the metadata reference](#metadata-reference) for a complete list of function properties.

Ensure your add-in only manifest file references your JSON file in the **\<Resources\>** section, similar to the following example.

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
  "allowCustomDataForDataTypeAny": true,
  "allowErrorForDataTypeAny": true,
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
      "id": "GETPLANETS", 
      "name": "GETPLANETS", 
      "description": "A function that uses the custom enum as a parameter.", 
      "parameters": [ 
        { 
          "name": "value", 
          "type": "string", 
          "customEnumType": "PLANETS" 
        }
      ]
    }
  ],
  "enums": [ 
    { 
      "id": "PLANETS", 
      "type": "string", 
      "values": [ 
        { 
          "name": "Mercury", 
          "value": "mercury", 
          "tooltip": "Mercury is the first planet from the sun." 
        }, 
        { 
          "name": "Venus", 
          "value": "venus", 
          "tooltip": "Venus is the second planet from the sun." 
        }
      ] 
    }
  ]
}
```

> [!NOTE]
> A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json) GitHub repository's commit history. As the project has been adjusted to automatically generate JSON, a full sample of handwritten JSON is only available in previous versions of the project.

## Metadata reference

### allowCustomDataForDataTypeAny

The `allowCustomDataForDataTypeAny` property is a Boolean data type. Setting this value to `true` allows a custom function to accept data types as parameters and return values. To learn more, see [Custom functions and data types](custom-functions-data-types-concepts.md).

> [!NOTE]
> Unlike most of the other JSON metadata properties, `allowCustomDataForDataTypeAny` is a top-level property and contains no sub-properties. See the preceding [JSON metadata code sample](#json-metadata-example) for an example of how to format this property.

If your custom function uses the `cellValueType` [parameter](#parameters), then setting the `allowCustomDataForDataTypeAny` isn't required to accept data types as parameters and return values.

### allowErrorForDataTypeAny

The `allowErrorForDataTypeAny` property is a Boolean data type. Setting the value to `true` allows a custom function to process errors as input values. All parameters with the type `any` or `any[][]` can accept errors as input values when `allowErrorForDataTypeAny` is set to `true`. The default `allowErrorForDataTypeAny` value is `false`.

> [!NOTE]
> Unlike the other JSON metadata properties, `allowErrorForDataTypeAny` is a top-level property and contains no sub-properties. See the preceding [JSON metadata code sample](#json-metadata-example) for an example of how to format this property.

### functions

The `functions` property is an array of custom function objects. The following table lists the properties of each object.

| Property      | Data type | Required | Description                                                                                                                                                                      |
| :------------ | :-------- | :------- | :------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `description` | string    | No       | The description of the function that end users see in Excel. For example, **Converts a Celsius value to Fahrenheit**.                                                            |
| `helpUrl`     | string    | No       | URL that provides information about the function. (It is displayed in a task pane.) For example, `http://contoso.com/help/convertcelsiustofahrenheit.html`.                      |
| `id`          | string    | Yes      | A unique ID for the function. This ID can only contain alphanumeric characters and periods and should not be changed after it is set.                                            |
| `name`        | string    | Yes      | The name of the function that end users see in Excel. In Excel, this function name is prefixed by the custom functions namespace that's specified in the add-in only manifest file. |
| `options`     | object    | No       | Enables you to customize some aspects of how and when Excel executes the function. See [options](#options) for details.                                                          |
| `parameters`  | array     | Yes      | Array that defines the input parameters for the function. See [parameters](#parameters) for details.                                                                             |
| `result`      | object    | Yes      | Object that defines the type of information that is returned by the function. See [result](#result) for details.                                                                 |

### enums

The `enums` property is an array of [enum](https://www.typescriptlang.org/docs/handbook/enums.html) objects. The following table lists the properties of each object.

> [!TIP]
> To learn about creating custom enums for your custom functions, see [Create custom enums for your custom functions](custom-functions-custom-enums.md). To learn about editing metadata for custom enums, see [Edit custom enums in JSON metadata](#edit-custom-enums-in-json-metadata).

| Property      | Data type | Required | Description                                                                                                                                                                      |
| :------------ | :-------- | :------- | :------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `name` | string    | Yes       | A brief description of the constant. |
| `tooltip` | string    | No       | Additional information about the constant that can be shown as a tooltip in user interfaces. |
| `value` | string    | Yes      | The value of the constant. |

### options

The `options` object enables you to customize some aspects of how and when Excel executes the function. The following table lists the properties of the `options` object.

| Property          | Data type | Required                               | Description |
| :---------------- | :-------- | :------------------------------------- | :---------- |
| `cancelable`      | Boolean   | No<br/><br/>Default value is `false`.  | If `true`, Excel calls the `CancelableInvocation` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function. Cancelable functions are typically only used for asynchronous functions that return a single result and need to handle the cancellation of a request for data. A function can't use both the `stream` and `cancelable` properties. |
| `capturesCallingObject` | Boolean | No<br/><br/>Default value is `false`. | If `true`, the data type being referenced by the custom function is passed as the first argument to the custom function. For more information, see [Reference the entity value as a calling object](excel-add-ins-dot-functions.md#reference-the-entity-value-as-a-calling-object). |
| `excludeFromAutoComplete` | Boolean | No<br/><br/>Default value is `false`. | If `true`, the custom function doesn't appear in the formula AutoComplete menu in Excel. For more information, see [Exclude custom functions from the Excel UI](excel-add-ins-dot-functions.md#exclude-custom-functions-from-the-excel-ui). A function can’t have both `excludeFromAutoComplete` and `linkedEntityLoadService` properties set to `true`. |
| `linkedEntityLoadService` | Boolean | No<br/><br/>Default value is `false`. | If `true`, the custom function provides a load service that returns up-to-date linked entity cell values for any linked entity IDs requested by Excel. A function can’t have both `excludeFromAutoComplete` and `linkedEntityLoadService` properties set to `true`. For more information, see [Linked entity load service function](excel-data-types-linked-entity-cell-values.md#linked-entity-load-service-function). |
| `requiresAddress` | Boolean   | No <br/><br/>Default value is `false`. | If `true`, your custom function can access the address of the cell that invoked it. The `address` property of the [invocation parameter](custom-functions-parameter-options.md#invocation-parameter) contains the address of the cell that invoked your custom function. A function can't use both the `stream` and `requiresAddress` properties. |
| `requiresParameterAddresses` | Boolean   | No <br/><br/>Default value is `false`. | If `true`, your custom function can access the addresses of the function's input parameters. This property must be used in combination with the `dimensionality` property of the [result](#result) object, and `dimensionality` must be set to `matrix`. See [Detect the address of a parameter](custom-functions-parameter-options.md#detect-the-address-of-a-parameter) for more information. |
| `requiresStreamAddress` | Boolean | No <br/><br/>Default value is `false`. | If `true`, the function can access the address of the cell calling the streaming function. The `address` property of the [invocation parameter](custom-functions-parameter-options.md#invocation-parameter) contains the address of the cell that invoked your streaming function. The function must also have `stream` set to `true`. |
| `requiresStreamParameterAddresses` | Boolean | No <br/><br/>Default value is `false`. | If `true`, the function can access the parameter addresses of the cell calling the streaming function. The `parameterAddresses` property of the [invocation parameter](custom-functions-parameter-options.md#invocation-parameter) contains the parameter addresses for your streaming function. The function must also have `stream` set to `true`. |
| `stream`          | Boolean   | No<br/><br/>Default value is `false`.  | If `true`, the function can output repeatedly to the cell even when invoked only once. This option is useful for rapidly-changing data sources, such as a stock price. The function should have no `return` statement. Instead, the result value is passed as the argument of the `StreamingInvocation.setResult` callback function. For more information, see [Make a streaming function](custom-functions-web-reqs.md#make-a-streaming-function). |
| `volatile`        | Boolean   | No <br/><br/>Default value is `false`. | If `true`, the function recalculates each time Excel recalculates, instead of only when the formula's dependent values have changed. A function can't use both the `stream` and `volatile` properties. If the `stream` and `volatile` properties are both set to `true`, the volatile property will be ignored. |

### parameters

The `parameters` property is an array of parameter objects. The following table lists the properties of each object.

|  Property  |  Data type  |  Required  |  Description  |
|:-----|:-----|:-----|:-----|
|  `description`  |  string  |  No |  A description of the parameter. This is displayed in Excel's IntelliSense.  |
|  `dimensionality`  |  string  |  No  |  Must be either `scalar` (a non-array value) or `matrix` (a 2-dimensional array).  |
|  `name`  |  string  |  Yes  |  The name of the parameter. This name is displayed in Excel's IntelliSense.  |
|  `type`  |  string  |  No  |  The data type of the parameter. Can be `boolean`, `number`, `string`, or `any`, which allows you to use of any of the previous three types. If this property is not specified, the data type defaults to `any`. |
| `cellValueType` | string | No | A subfield of the `type` property. Specifies the Excel data types accepted by the custom function. Accepts the case-insensitive values `cellvalue`, `booleancellvalue`, `doublecellvalue`, `entitycellvalue`, `errorcellvalue`, `linkedentitycellvalue`, `localimagecellvalue`, `stringcellvalue`, and `webimagecellvalue`. <br/><br/>The `type` field must have the value `any` to use the `cellValueType` subfield. |
| `customEnumType` | string | No | The `id` of the enum in the `enums` array. This associates the custom enum with the function and enables Excel to display the enum members in the formula AutoComplete menu. |
|  `optional`  | Boolean | No | If `true`, the parameter is optional. |
|`repeating`| Boolean | No | If `true`, parameters populate from a specified array. Note that all repeating parameters are considered optional parameters by definition.  |

> [!TIP]
> See the following code snippet for an example of how to format the `cellValueType` parameter in JSON metadata.
>
> ```json
> "parameters": [
>     {
>         "name": "range",
>         "description": "the input range",
>         "type": "any",
>             "cellValueType": "webimagecellvalue"
>     }
> ]
> ```

### result

The `result` object defines the type of information that is returned by the function. The following table lists the properties of the `result` object.

| Property         | Data type | Required | Description                                                                          |
| :--------------- | :-------- | :------- | :----------------------------------------------------------------------------------- |
| `dimensionality` | string    | No       | Must be either `scalar` (a non-array value) or `matrix` (a 2-dimensional array). |
| `type` | string    | No       | The data type of the result. Can be `boolean`, `number`, `string`, or `any` (which allows you to use of any of the previous three types). If this property is not specified, the data type defaults to `any`. |

## Associate function names with JSON metadata

For a function to work properly, you need to associate the function's `id` property with the JavaScript implementation. Make sure there is an association, otherwise the function won't be registered and isn't useable in Excel. The following code sample shows how to make the association using the `CustomFunctions.associate()` function. The sample defines the custom function `add` and associates it with the object in the JSON metadata file where the value of the `id` property is **ADD**.

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

The following sample shows the JSON metadata that corresponds to the functions defined in the preceding JavaScript code sample. The `id` and `name` property values are in uppercase, which is a best practice when describing your custom functions. You only need to add this JSON if you are preparing your own JSON file manually and not using autogeneration. For more information on autogeneration, see [Autogenerate JSON metadata for custom functions](custom-functions-json-autogeneration.md).

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/office-js/custom-functions.schema.json",
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

## Edit custom enums in JSON metadata

Create or edit enum metadata directly with the `enums` property. Each custom enum must have a unique ID value and type value of either `string` or `number`. Mixed type enums are not supported.

If you manually create the JSON metadata for your custom enum, you can associate those enums with either TypeScript or JavaScript custom functions. To learn more about creating custom enums, see [Create custom enums for your custom functions](custom-functions-custom-enums.md).

The following JSON snippet shows the metadata for two enums: a `PLANETS` enum  that contains the planets Mercury and Venus, and a `DAYS` enum that includes the days Monday and Tuesday.

```json
"enums": [ 
  { 
    "id": "PLANETS", 
    "type": "string", 
    "values": [ 
      { 
        "name": "Mercury", 
        "value": "mercury", 
        "tooltip": "Mercury is the first planet from the sun." 
      }, 
      { 
        "name": "Venus", 
        "value": "venus", 
        "tooltip": "Venus is the second planet from the sun." 
      }
    ] 
  },
  {
    "id": "DAYS", 
    "type": "number", 
    "values": [ 
      { 
        "name": "Monday",
        "value": 1,
        "tooltip": "Monday is the first working day of a week."
      },
      { 
        "name": "Tuesday",
        "value": 2,
        "tooltip": "Tuesday is the second working day of a week."
      }
    ] 
  }
]
```

Each constant in the `values` array of the enum is an object with the following properties.

- **value**: The value of the constant.
- **name**: A brief description of the constant.
- **tooltip** (Optional): Additional information about the constant that can be shown as a tooltip in user interfaces.

To associate the custom enum with a function, add the property `customEnumType` to the `parameters` object. The `customEnumType` value should match the `id` of the enum. Note that the `customEnumType` value is not case-sensitive. The following JSON snippet shows a `functions` object associated with the `PLANETS` enum.

```json
"functions": [ 
  {
    "description": "A function that uses the custom enum as a parameter.", 
    "id": "GETPLANETS", 
    "name": "GETPLANETS", 
    "parameters": [ 
      { 
        "name": "value", 
        "type": "string", 
        "customEnumType": "PLANETS" 
      }
    ], 
    "result": {} 
  } 
]
```

## Next steps

Learn the [best practices for naming your function](custom-functions-naming.md) or discover how to [localize your function](custom-functions-naming.md#localize-custom-functions) using the previously described handwritten JSON method.

## See also

- [Autogenerate JSON metadata for custom functions](custom-functions-json-autogeneration.md)
- [Custom functions parameter options](custom-functions-parameter-options.md)
- [Create custom functions in Excel](custom-functions-overview.md)
