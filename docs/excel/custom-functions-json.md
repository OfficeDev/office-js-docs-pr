---
ms.date: 09/20/2018
description: Define metadata for custom functions in Excel.
title: Metadata for custom functions in Excel
---

# Custom functions metadata

When you include [custom functions](custom-functions-overview.md) in an Excel add-in, you must host a JSON file that contains metadata about the functions (in addition to hosting a JavaScript file with the functions and a UI-less HTML file to serve as the parent of the JavaScript file). This article describes the format of the JSON file with examples.

A complete sample JSON file is available [here](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).

## Functions array

The metadata is a JSON object that contains a single `functions` property whose value is an array of objects. Each of these objects represents one custom function. The following table contains its properties:

|  Property  |  Data Type  |  Required?  |  Description  |
|:-----|:-----|:-----|:-----|
|  `description`  |  string  |  No  |  A description of the function that appears in the Excel UI. For example, "Converts a Celsius value to Fahrenheit". |
|  `helpUrl`  |  string  |   No  |  URL where your users can get help about the function. (It is displayed in a taskpane.) For example, "http://contoso.com/help/convertcelsiustofahrenheit.html"  |
|  `name`  |  string  |  Yes  |  The name of the function as it will appear (prepended with a namespace) in the Excel UI when a user is selecting a function. It should be the same as the function's name where it is defined in the JavaScript. |
|  `options`  |  object  |  No  |  Configure how Excel processes the function. See [options object](#options-object) for details. |
|  `parameters`  |  array  |  Yes  |  Metadata about the parameters to the function. See [parameters array](#parameters-array)  for details. |
|  `result`  |  object  |  Yes  |  Metadata about the value returned by the function. See [result object](#result-object) for details. |

## Options object

The `options` object configures how Excel processes the function. The following table contains its properties:

|  Property  |  Data Type  |  Required?  |  Description  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  boolean  |  No, default is `false`.  |  If `true`, Excel calls the `onCanceled` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function. If you use this option, Excel will call the JavaScript function with an additional `caller` parameter. (Do ***not*** register this parameter in the `parameters` property). In the body of the function, a handler must be assigned to the `caller.onCanceled` member.|
|  `stream`  |  boolean  |  No, default is `false`.  |  If `true`, the function can output repeatedly to the cell even when invoked only once. This option is useful for rapidly-changing data sources, such as a stock price. If you use this option, Excel will call the JavaScript function with an additional `caller` parameter. (Do ***not*** register this parameter in the `parameters` property). The function should have no `return` statement. Instead, the result value is passed as the argument of the `caller.setResult` callback method.|

## Parameters array

The `parameters` property is an array of objects. Each of these objects represents a parameter. The following table contains its properties:

|  Property  |  Data Type  |  Required?  |  Description  |
|:-----|:-----|:-----|:-----|
|  `description`  |  string  |  No |  A description of the parameter.  |
|  `dimensionality`  |  string  |  Yes  |  Must be either "scalar", meaning a non-array value, or "matrix", meaning an array of row arrays.  |
|  `name`  |  string  |  Yes  |  The name of the parameter. This name is displayed in Excel's IntelliSense.  |
|  `type`  |  string  |  Yes  |  The data type of the parameter. Must be "boolean", "number", or "string".  |

## Result object

The `results` property provides metadata about the value returned from the function. The following table contains its properties:

|  Property  |  Data Type  |  Required?  |  Description  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  string  |  No  |  Must be either "scalar", meaning a non-array value, or "matrix", meaning an array of row arrays.  |
|  `type`  |  string  |  Yes  |  The data type of the parameter. Must be "boolean", "number", or "string".  |

## Example

The following JSON code is an example of a metadata file for custom functions.

```json
{
	"functions": [
		{
			"name": "ADD42", 
			"description":  "Adds 42 to the input number",
			"helpUrl": "http://dev.office.com",
			"result": {
				"type": "number",
				"dimensionality": "scalar"
			},
			"parameters": [
				{
					"name": "num",
					"description": "Number",
					"type": "number",
					"dimensionality": "scalar"
				}
			]
		},
		{
			"name": "ADD42ASYNC", 
			"description":  "asynchronously wait 250ms, then add 42",
			"helpUrl": "http://dev.office.com",
			"result": {
				"type": "number",
				"dimensionality": "scalar"
			},
			"parameters": [
				{
					"name": "num",
					"description": "Number",
					"type": "number",
					"dimensionality": "scalar"
				}
			]
		},
		{
			"name": "ISEVEN", 
			"description":  "Determines whether a number is even",
			"helpUrl": "http://dev.office.com",
			"result": {
				"type": "boolean",
				"dimensionality": "scalar"
			},
			"parameters": [
				{
					"name": "num",
					"description": "the number to be evaluated",
					"type": "number",
					"dimensionality": "scalar"
				}
			]
		},
		{
			"name": "GETDAY",
			"description": "Gets the day of the week",
			"helpUrl": "http://dev.office.com",
			"result": {
				"type": "string"
			},
			"parameters": []
		},
		{
			"name": "INCREMENTVALUE", 
			"description":  "Counts up from zero",
			"helpUrl": "http://dev.office.com",
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
			"name": "SECONDHIGHEST", 
			"description":  "gets the second highest number from a range",
			"helpUrl": "http://dev.office.com",
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

## See also

* [Create custom functions in Excel](custom-functions-overview.md)
* [Runtime for Excel custom functions](custom-functions-runtime.md)
* [Custom functions best practices](custom-functions-best-practices.md)