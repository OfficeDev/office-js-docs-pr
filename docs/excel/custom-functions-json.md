---
ms.date: 09/27/2018
description: Define metadata for custom functions in Excel.
title: Metadata for custom functions in Excel
---

# Custom functions metadata (preview)

When you define [custom functions](custom-functions-overview.md) within your Excel add-in, your add-in project must include a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end users. This article describes the format of the JSON metadata file.

For information about the other files that you must include in your add-in project to enable custom functions, see [Create custom functions in Excel](custom-functions-overview.md).

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

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
> A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json) GitHub repository.

## functions 

The `functions` property is an array of custom function objects. The following table lists the properties of each object.

|  Property  |  Data type  |  Required  |  Description  |
|:-----|:-----|:-----|:-----|
|  `description`  |  string  |  No  |  The description of the function that end users see in Excel. For example, **Converts a Celsius value to Fahrenheit**. |
|  `helpUrl`  |  string  |   No  |  URL that provides information about the function. (It is displayed in a task pane.) For example, **http://contoso.com/help/convertcelsiustofahrenheit.html**. |
| `id`     | string | Yes | A unique ID for the function. This ID should not be changed after it is set. |
|  `name`  |  string  |  Yes  |  The name of the function that end users see in Excel. In Excel, this function name will be prefixed by the custom functions namespace that's specified in the XML manifest file. |
|  `options`  |  object  |  No  |  Enables you to customize some aspects of how and when Excel executes the function. See [options object](#options-object) for details. |
|  `parameters`  |  array  |  Yes  |  Array that defines the input parameters for the function. See [parameters array](#parameters-array)  for details. |
|  `result`  |  object  |  Yes  |  Object that defines the type of information that is returned by the function. See [result object](#result-object) for details. |

## options

The `options` object enables you to customize some aspects of how and when Excel executes the function. The following table lists the properties of the `options` object.

|  Property  |  Data type  |  Required  |  Description  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  boolean  |  No<br/><br/>Default value is `false`.  |  If `true`, Excel calls the `onCanceled` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function. If you use this option, Excel will call the JavaScript function with an additional `caller` parameter. (Do ***not*** register this parameter in the `parameters` property). In the body of the function, a handler must be assigned to the `caller.onCanceled` member. For more information, see [Canceling a function](custom-functions-overview.md#canceling-a-function). |
|  `stream`  |  boolean  |  No<br/><br/>Default value is `false`.  |  If `true`, the function can output repeatedly to the cell even when invoked only once. This option is useful for rapidly-changing data sources, such as a stock price. If you use this option, Excel will call the JavaScript function with an additional `caller` parameter. (Do ***not*** register this parameter in the `parameters` property). The function should have no `return` statement. Instead, the result value is passed as the argument of the `caller.setResult` callback method. For more information, see [Streaming functions](custom-functions-overview.md#streaming-functions). |

## parameters

The `parameters` property is an array of parameter objects. The following table lists the properties of each object.

|  Property  |  Data type  |  Required  |  Description  |
|:-----|:-----|:-----|:-----|
|  `description`  |  string  |  No |  A description of the parameter.  |
|  `dimensionality`  |  string  |  No  |  Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).  |
|  `name`  |  string  |  Yes  |  The name of the parameter. This name is displayed in Excel's intelliSense.  |
|  `type`  |  string  |  No  |  The data type of the parameter. Must be **boolean**, **number**, or **string**.  |

## result

The `results` object defines the type of information that is returned by the function. The following table lists the properties of the `result` object.

|  Property  |  Data type  |  Required  |  Description  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  string  |  No  |  Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array). |
|  `type`  |  string  |  Yes  |  The data type of the parameter. Must be **boolean**, **number**, or **string**.  |

## See also

* [Create custom functions in Excel](custom-functions-overview.md)
* [Runtime for Excel custom functions](custom-functions-runtime.md)
* [Custom functions best practices](custom-functions-best-practices.md)
* [Excel custom functions tutorial](excel-tutorial-custom-functions.md)