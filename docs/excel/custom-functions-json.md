---
ms.date: 03/29/2019
description: Define metadata for custom functions in Excel.
title: Metadata for custom functions in Excel (preview)
localization_priority: Normal
---

# Custom functions metadata (preview)

When you define [custom functions](custom-functions-overview.md) within your Excel add-in, your add-in project includes a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end users. This file is generated either:

- by you, in a handwritten JSON file
- from the JSDoc comments you enter at the beginning of your function

Custom functions are registered when the user runs the add-in for the first time and after that are available to the same user in all workbooks.

This article describes the format of the JSON metadata file, assuming you are writing it by hand. For information about JSDoc comment JSON file generation, see [Generate JSON metadata for custom functions](custom-functions-json-autogeneration.md).

For information about the other files that you must include in your add-in project to enable custom functions, see [Create custom functions in Excel](custom-functions-overview.md).

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

> Server settings on the server that hosts the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel Online.

## Example metadata

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
		"type": "string"
	  },
	  "parameters": []
	},
	{
	  "id": "INCREMENTVALUE",
	  "name": "INCREMENTVALUE", 
	  "description":  "Count up from zero",
	  "helpUrl": "http://www.contoso.com/help",
	  "result": {
		"type": "number",
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
	  "description":  "Get the second highest number from a range",
	  "helpUrl": "http://www.contoso.com/help",
	  "result": {
		"type": "number",
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
> A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/src/functions/functions.json) GitHub repository.

## functions 

The `functions` property is an array of custom function objects. The following table lists the properties of each object.

|  Property  |  Data type  |  Required  |  Description  |
|:-----|:-----|:-----|:-----|
|  `description`  |  string  |  No  |  The description of the function that end users see in Excel. For example, **Converts a Celsius value to Fahrenheit**. |
|  `helpUrl`  |  string  |   No  |  URL that provides information about the function. (It is displayed in a task pane.) For example, **http://contoso.com/help/convertcelsiustofahrenheit.html**. |
| `id`     | string | Yes | A unique ID for the function. This ID can only contain alphanumeric characters and periods and should not be changed after it is set. |
|  `name`  |  string  |  Yes  |  The name of the function that end users see in Excel. In Excel, this function name will be prefixed by the custom functions namespace that's specified in the XML manifest file. |
|  `options`  |  object  |  No  |  Enables you to customize some aspects of how and when Excel executes the function. See [options](#options) for details. |
|  `parameters`  |  array  |  Yes  |  Array that defines the input parameters for the function. See [parameters](#parameters)  for details. |
|  `result`  |  object  |  Yes  |  Object that defines the type of information that is returned by the function. See [result](#result) for details. |

## options

The `options` object enables you to customize some aspects of how and when Excel executes the function. The following table lists the properties of the `options` object.

|  Property  |  Data type  |  Required  |  Description  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  boolean  |  No<br/><br/>Default value is `false`.  |  If `true`, Excel calls the `onCanceled` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function. If you use this option, Excel will call the JavaScript function with an additional `caller` parameter. (Do ***not*** register this parameter in the `parameters` property). In the body of the function, a handler must be assigned to the `caller.onCanceled` member. For more information, see [Canceling a function](custom-functions-web-reqs.md#canceling-a-function). |
|  `requiresAddress`  | boolean | No <br/><br/>Default value is `false`. | <br /><br /> If true, your custom function can access the address of the cell that invoked your custom function. To get the address of the cell that invoked your custom function, use context.address in your custom function. For more information, see [Determine which cell invoked your custom function](/office/dev/add-ins/excel/custom-functions-overview#determine-which-cell-invoked-your-custom-function). Custom functions cannot be set as both streaming and requiresAddress. When using this option, the 'invocationContext' parameter must be the last parameter passed in options. |
|  `stream`  |  boolean  |  No<br/><br/>Default value is `false`.  |  If `true`, the function can output repeatedly to the cell even when invoked only once. This option is useful for rapidly-changing data sources, such as a stock price. If you use this option, Excel will call the JavaScript function with an additional `caller` parameter. (Do ***not*** register this parameter in the `parameters` property). The function should have no `return` statement. Instead, the result value is passed as the argument of the `caller.setResult` callback method. For more information, see [Streaming functions](custom-functions-web-reqs.md#streaming-functions). |
|  `volatile`  | boolean | No <br/><br/>Default value is `false`. | <br /><br /> If `true`, the function will recalculate each time Excel recalculates, instead of only when the formula's dependent values have changed. A function cannot be both streaming and volatile. If the `stream` and `volatile` properties are both set to `true`, the volatile option will be ignored. |

## parameters

The `parameters` property is an array of parameter objects. The following table lists the properties of each object.

|  Property  |  Data type  |  Required  |  Description  |
|:-----|:-----|:-----|:-----|
|  `description`  |  string  |  No |  A description of the parameter. This is displayed in Excel's intelliSense.  |
|  `dimensionality`  |  string  |  No  |  Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).  |
|  `name`  |  string  |  Yes  |  The name of the parameter. This name is displayed in Excel's intelliSense.  |
|  `type`  |  string  |  No  |  The data type of the parameter. Can be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types. If this property is not specified, the data type defaults to **any**. |
|  `optional`  | boolean | No | If `true`, the parameter is optional. |

>[!NOTE]
> If the `type` property of an optional parameter is either not specified or set to `any`, you may notice issues such as linting errors in your IDE and optional parameters not being displayed when the function is being entered into a cell in Excel. This is projected to change in December of 2018.

## result

The `result` object defines the type of information that is returned by the function. The following table lists the properties of the `result` object.

|  Property  |  Data type  |  Required  |  Description  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  string  |  No  |  Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array). |
|  `type`  |  string  |  Yes  |  The data type of the parameter. Must be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types. |

## See also

* [Create custom functions in Excel](custom-functions-overview.md)
* [Runtime for Excel custom functions](custom-functions-runtime.md)
* [Custom functions best practices](custom-functions-best-practices.md)
* [Custom functions changelog](custom-functions-changelog.md)
* [Excel custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md)
