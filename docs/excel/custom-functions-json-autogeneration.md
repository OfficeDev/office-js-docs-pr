---
title: Autogenerate JSON metadata for custom functions
description: Use JSDoc tags to dynamically create your custom functions JSON metadata.
ms.date: 12/09/2025
ms.localizationpriority: medium
---

# Autogenerate JSON metadata for custom functions

When an Excel custom function is written in JavaScript or TypeScript, [JSDoc tags](https://jsdoc.app/) are used to provide extra information about the custom function. We provide a [Webpack](https://webpack.js.org/) plugin that uses these JSDoc tags to automatically create the JSON metadata file at build time. Using the plugin saves you from the effort of [manually editing the JSON metadata file](custom-functions-json.md).

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## CustomFunctionsMetadataPlugin

The plugin is [CustomFunctionsMetadataPlugin](https://github.com/OfficeDev/Office-Addin-Scripts/blob/master/packages/custom-functions-metadata-plugin/README.md). To install and configure it, use the following steps.

> [!NOTE]
> 
> - The tool can be used only in a NodeJS-based project.
> - These instructions assume that your project uses [Webpack](https://webpack.js.org/) and that you have it installed and configured.
> - If your custom function add-in project is created with the [Yeoman generator for Office Add-ins](../develop/yeoman-generator-overview.md), Webpack is installed and all of these steps are done automatically, but when applicable, you must do the steps in [Multiple custom function source files](#multiple-custom-function-source-files) manually.

1. Open a Command Prompt or bash shell and, in the root of the project, run `npm install custom-functions-metadata-plugin`.
1. Open the webpack.config.js file and add the following line at the top: `const CustomFunctionsMetadataPlugin = require("custom-functions-metadata-plugin");`.
1. Scroll down to the `plugins` array and add the following to the top of the array. Change the `input` path and filename as needed to match your project, but the `output` value must be `"functions.json"`. If you're using TypeScript, use the \*.ts source file name, *not* the transpiled \*.js file.

   ```js
   new CustomFunctionsMetadataPlugin({
      output: "functions.json",
      input: "./src/functions/functions.js", 
   }),
   ```

### Multiple custom function source files

If, and only if, you have organized your custom functions into multiple source files, there are additional steps.

1. In the webpack.config.js file, replace the string value of `input` with an array of string URLs that point to each of the files. The following is an example:

   ```js
   new CustomFunctionsMetadataPlugin({
      output: "functions.json",
      input: [
               "./src/functions/someFunctions.js", 
               "./src/functions/otherFunctions.js"
             ], 
   }),
   ```

1. Scroll to the `entry.functions` property, and replace its value with the same array you used in the preceding step. The following is an example:

   ```js
   entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      taskpane: ["./src/taskpane/taskpane.js", "./src/taskpane/taskpane.html"],
      functions: [
               "./src/functions/someFunctions.js", 
               "./src/functions/otherFunctions.js"
             ],
    },
   ```

## Run the tool

You don't have to do anything to run the tool. When Webpack runs, it creates the functions.json file and puts it in memory in development mode, or in the /dist folder in production mode.

## Basics of JSDoc tags

Add the `@customfunction` tag in the code comments for a JavaScript or TypeScript function to mark it as a custom function.

The function parameter types may be provided using the [@param](#param) tag in JavaScript, or from the [Function type](https://www.typescriptlang.org/docs/handbook/functions.html) in TypeScript. For more information, see the [@param](#param) tag and [Types](#types) sections.

### Add a description to a function

The description is displayed to the user as help text when they need help to understand what your custom function does. The description doesn't require any specific tag. Just enter a short text description in the JSDoc comment. In general the description is placed at the start of the JSDoc comment section, but it will work no matter where it is placed.

To see examples of the built-in function descriptions, open Excel, go to the **Formulas** tab, and choose **Insert function**. You can then browse through all the function descriptions, and also see your own custom functions listed.

In the following example, the phrase "Calculates the volume of a sphere." is the description for the custom function.

```js
/**
/* Calculates the volume of a sphere.
/* @customfunction VOLUME
...
 */
```

## Supported JSDoc tags

The following JSDoc tags are supported in Excel custom functions.

- [@cancelable](#cancelable)
- [@capturesCallingObject](#capturesCallingObject)
- [@customenum](#customenum) *{type}*
- [@customfunction](#customfunction) *id* *name*
- [@excludeFromAutoComplete](#excludeFromAutoComplete)
- [@helpurl](#helpurl) *url*
- [@linkedEntityLoadService](#linkedEntityLoadService)
- [@param](#param) *{type}* *name* *description*
- [@requiresAddress](#requiresAddress)
- [@requiresParameterAddresses](#requiresParameterAddresses)
- [@returns](#returns) *{type}*
- [@streaming](#streaming)
- [@supportSync](#supportsync)
- [@volatile](#volatile)

---
<a id="cancelable"></a>

### @cancelable

Indicates that a custom function performs an action when the function is canceled.

The last function parameter must be of type `CustomFunctions.CancelableInvocation`. The function can assign a function to the `oncanceled` property to denote the result when the function is canceled.

If the last function parameter is of type `CustomFunctions.CancelableInvocation`, it will be considered `@cancelable` even if the tag isn't present.

A function can't have both `@cancelable` and `@streaming` tags.

<a id="capturesCallingObject"></a>

### @capturesCallingObject

This tag works with Excel [data types](excel-data-types-overview.md). It specifies that the data type being referenced by the custom function is passed as the first argument to the custom function. For more information, see [Reference the entity value as a calling object](excel-add-ins-dot-functions.md#reference-the-entity-value-as-a-calling-object).

<a id="customenum"></a>

### @customenum

Syntax: @customenum *{type}*

This tag indicates that a set of values is a custom enum. For more information, see [Create custom enums for your custom functions](custom-functions-custom-enums.md).

<a id="customfunction"></a>

### @customfunction

Syntax: @customfunction *id* *name*

This tag indicates that the JavaScript/TypeScript function is an Excel custom function. It is required to create metadata for the custom function.

The following shows an example of this tag.

```js
/**
 * Increments a value once a second.
 * @customfunction
 * ...
 */
```

#### id

The `id` identifies a custom function.

- If `id` isn't provided, the JavaScript/TypeScript function name is converted to uppercase and disallowed characters are removed.
- The `id` must be unique for all custom functions.
- The allowed characters are limited to: A-Z, a-z, 0-9, underscores (\_), and period (.).

In the following example, increment is the `id` and the `name` of the function.

```js
/**
 * Increments a value once a second.
 * @customfunction INCREMENT
 * ...
 */
```

#### name

Provides the display `name` for the custom function.

- If name isn't provided, the id is also used as the name.
- Allowed characters: Letters [Unicode Alphabetic character](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), numbers, period (.), and underscore (\_).
- Must start with a letter.
- Maximum length is 128 characters.

In the following example, INC is the `id` of the function and `increment` is the `name`.

```js
/**
 * Increments a value once a second.
 * @customfunction INC INCREMENT
 * ...
 */
```

### description

A description appears to users in Excel as they are entering the function and specifies what the function does. A description doesn't require any specific tag. Add a description to a custom function by adding a phrase to describe what the function does inside the JSDoc comment. By default, whatever text is untagged in the JSDoc comment section will be the description of the function.

In the following example, the phrase "A function that adds two numbers" is the description for the custom function with the id property of `ADD`.

```js
/**
 * A function that adds two numbers.
 * @customfunction ADD
 * ...
 */
```

<a id="excludeFromAutoComplete"></a>

### @excludeFromAutoComplete

The `@excludeFromAutoComplete` tag ensures that the custom function doesn't appear in the formula AutoComplete menu in Excel. For more information, see [Exclude custom functions from the Excel UI](excel-add-ins-dot-functions.md#exclude-custom-functions-from-the-excel-ui).

A function can’t have both `@excludeFromAutoComplete` and `@linkedEntityLoadService` tags.

<a id="helpurl"></a>

### @helpurl

Syntax: @helpurl *url*

The provided *url* is displayed in Excel.

In the following example, the `helpurl` is `http://www.contoso.com/weatherhelp`.

```js
/**
 * A function which streams the temperature in a town you specify.
 * @customfunction getTemperature
 * @helpurl http://www.contoso.com/weatherhelp
 * ...
 */
```

<a id="linkedEntityLoadService"></a>

### @linkedEntityLoadService

The `@linkedEntityLoadService` tag designates that the function is a linked entity load service that returns [linked entity cell values](excel-data-types-linked-entity-cell-values.md) for linked entity IDs requested by Excel.

A function can’t have both `@excludeFromAutoComplete` and `@linkedEntityLoadService` tags.

<a id="param"></a>

### @param

#### JavaScript

JavaScript Syntax: @param {type} *name* *description*

- `{type}` specifies the type info within curly braces. See the [Types](#types) section for more information about the types which may be used. If no type is specified, the default type `any` will be used.
- `name` specifies the parameter that the @param tag applies to. It is required.
- `description` provides the description which appears in Excel for the function parameter. It is optional.

To denote a custom function parameter as optional, put square brackets around the parameter name. For example, `@param {string} [text] Optional text`.

> [!NOTE]
> The default value for optional parameters is `null`.

The following example shows an ADD function that adds two or three numbers, with the third number as an optional parameter.

```js
/**
 * A function which sums two, or optionally three, numbers.
 * @customfunction ADDNUMBERS
 * @param firstNumber {number} First number to add.
 * @param secondNumber {number} Second number to add.
 * @param [thirdNumber] {number} Optional third number you wish to add.
 * ...
 */
```

#### TypeScript

TypeScript Syntax: @param *name* *description*

- `name` specifies the parameter that the @param tag applies to. It is required.
- `description` provides the description which appears in Excel for the function parameter. It is optional.

See the [Types](#types) section for more information about the function parameter types which may be used.

To denote a custom function parameter as optional, do one of the following:

- Use an optional parameter. For example: `function f(text?: string)`
- Give the parameter a default value. For example: `function f(text: string = "abc")`

For detailed description of the @param see: [JSDoc](https://jsdoc.app/tags-param.html)

> [!NOTE]
> The default value for optional parameters is `null`.

The following example shows the `add` function that adds two numbers.

```ts
/**
 * Adds two numbers.
 * @customfunction 
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
function add(first: number, second: number): number {
  return first + second;
}
```

<a id="requiresAddress"></a>

### @requiresAddress

Indicates that the address of the cell where the function is being evaluated should be provided.

The last function parameter must be of type `CustomFunctions.Invocation` or a derived type to use `@requiresAddress`. When the function is called, the `address` property will contain the address.

The following sample shows how to use the `invocation` parameter in combination with `@requiresAddress` to return the address of the cell that invoked your custom function. See [Invocation parameter](custom-functions-parameter-options.md#invocation-parameter) for more information.

```js
/**
 * Return the address of the cell that invoked the custom function. 
 * @customfunction
 * @param {number} first First parameter.
 * @param {number} second Second parameter.
 * @param {CustomFunctions.Invocation} invocation Invocation object. 
 * @requiresAddress 
 */
function getAddress(first, second, invocation) {
  const address = invocation.address;
  return address;
}
```

<a id="requiresParameterAddresses"></a>

### @requiresParameterAddresses

Indicates that the function should return the addresses of input parameters.

The last function parameter must be of type `CustomFunctions.Invocation` or a derived type to use  `@requiresParameterAddresses`. The JSDoc comment must also include an `@returns` tag specifying that the return value be a matrix, such as `@returns {string[][]}` or `@returns {number[][]}`. See [Matrix types](#matrix-type) for additional information.

When the function is called, the `parameterAddresses` property will contain the addresses of the input parameters.

The following sample shows how to use the `invocation` parameter in combination with `@requiresParameterAddresses` to return the addresses of three input parameters. See [Detect the address of a parameter](custom-functions-parameter-options.md#detect-the-address-of-a-parameter) for more information.

```js
/**
 * Return the addresses of three parameters. 
 * @customfunction
 * @param {string} firstParameter First parameter.
 * @param {string} secondParameter Second parameter.
 * @param {string} thirdParameter Third parameter.
 * @param {CustomFunctions.Invocation} invocation Invocation object. 
 * @returns {string[][]} The addresses of the parameters, as a 2-dimensional array.
 * @requiresParameterAddresses
 */
function getParameterAddresses(firstParameter, secondParameter, thirdParameter, invocation) {
  const addresses = [
    [invocation.parameterAddresses[0]],
    [invocation.parameterAddresses[1]],
    [invocation.parameterAddresses[2]]
  ];
  return addresses;
}
```

<a id="returns"></a>

### @returns

Syntax: @returns {*type*}

Provides the type for the return value.

If `{type}` is omitted, the TypeScript type info will be used. If there is no type info, the type will be `any`.

The following example shows the `add` function that uses the `@returns` tag.

```ts
/**
 * Adds two numbers.
 * @customfunction 
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
function add(first: number, second: number): number {
  return first + second;
}
```

<a id="streaming"></a>

### @streaming

Indicates that a custom function is a streaming function.

The last parameter is of type `CustomFunctions.StreamingInvocation<ResultType>`.
The function returns `void`.

Streaming functions don't return values directly, instead they call `setResult(result: ResultType)` using the last parameter.

Exceptions thrown by a streaming function are ignored. `setResult()` may be called with Error to indicate an error result. For an example of a streaming function and more information, see [Make a streaming function](custom-functions-web-reqs.md#make-a-streaming-function).

Streaming functions can't be marked as [@supportSync](#supportsync) or [@volatile](#volatile).

<a id="supportsync"></a>

### @supportSync

Indicates that a custom function supports synchronous processes in Excel, like evaluate and conditional format actions.

If the function uses `Excel.RequestContext`, call the `setInvocation` method of `Excel.RequestContext` and pass in the `CustomFunctions.Invocation` object. For more information, see [Synchronous custom functions](custom-functions-synchronous.md).

Synchronous custom functions can't be marked as [@streaming](#streaming) or [@volatile](#volatile). If `@streaming` or `@volatile` tags are used, the `@supportSync` tag is ignored.

<a id="volatile"></a>

### @volatile

A volatile function is one whose result isn't the same from one moment to the next, even if it takes no arguments or the arguments haven't changed. Excel re-evaluates cells that contain volatile functions, together with all dependents, every time that a calculation is done. For this reason, too much reliance on volatile functions can make recalculation times slow, so use them sparingly.

Volatile functions can't be marked as [@streaming](#streaming) or [@supportSync](#supportsync).

The following function is volatile and uses the `@volatile` tag.

```js
/**
 * Simulates rolling a 6-sided die.
 * @customfunction
 * @volatile
 */
function roll6sided(): number {
  return Math.floor(Math.random() * 6) + 1;
}
```

---

## Types

By specifying a parameter type, Excel will convert values into that type before calling the function. If the type is `any`, no conversion will be performed.

### Value types

A single value may be represented using one of the following types: `boolean`, `number`, `string`.

### Cell value type

Use the `type` subfield `cellValueType` to specify that a custom function accept and return Excel data types. The `type` value must be `any` to use the `cellValueType` subfield. Accepted `cellValueType` values are:

- `Excel.CellValue`
- `Excel.BooleanCellValue`
- `Excel.DoubleCellValue`
- `Excel.EntityCellValue`
- `Excel.ErrorCellValue`
- `Excel.LinkedEntityCellValue`
- `Excel.LocalImageCellValue`
- `Excel.StringCellValue`
- `Excel.WebImageCellValue`

For a code sample using the `Excel.EntityCellValue` type, see [Input an entity value](custom-functions-data-types-concepts.md#input-an-entity-value).

### Matrix type

Use a two-dimensional array type to have the parameter or return value be a matrix of values. For example, the type `number[][]` indicates a matrix of numbers and `string[][]` indicates a matrix of strings.

### Error type

A non-streaming function can indicate an error by returning an Error type.

A streaming function can indicate an error by calling `setResult()` with an Error type.

### Promise

A custom function can return a promise that provides the value when the promise is resolved. If the promise is rejected, then the custom function will throw an error.

### Other types

Any other type will be treated as an error.

## Next steps

Learn about [naming and localization for custom functions](custom-functions-naming.md).

## See also

- [Manually create JSON metadata for custom functions](custom-functions-json.md)
- [Create custom functions in Excel](custom-functions-overview.md)
