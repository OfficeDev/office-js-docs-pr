---
ms.date: 07/15/2019
description: Use JSDoc tags to dynamically create your custom functions JSON metadata.
title: Autogenerate JSON metadata for custom functions
localization_priority: Priority
---

# Autogenerate JSON metadata for custom functions

When an Excel custom function is written in JavaScript or TypeScript, [JSDoc tags](https://jsdoc.app/) are used to provide extra information about the custom function. The JSDoc tags are then used at build time to create the [JSON metadata file](custom-functions-json.md). Using JSDoc tags saves you from the effort of manually editing the JSON metadata file.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Add the `@customfunction` tag in the code comments for a JavaScript or TypeScript function to mark it as a custom function.

The function parameter types may be provided using the [@param](#param) tag in JavaScript, or from the [Function type](https://www.typescriptlang.org/docs/handbook/functions.html) in TypeScript. For more information, see the [@param](#param) tag and [Types](#types) sections.

### Adding a description to a function

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


## JSDoc Tags
The following JSDoc tags are supported in Excel custom functions:
* [@cancelable](#cancelable)
* [@customfunction](#customfunction) id name
* [@helpurl](#helpurl) url
* [@param](#param) _{type}_ name description
* [@requiresAddress](#requiresAddress)
* [@returns](#returns) _{type}_
* [@streaming](#streaming)
* [@volatile](#volatile)

---
### @cancelable
<a id="cancelable"/>

Indicates that a custom function wants to perform an action when the function is canceled.

The last function parameter must be of type `CustomFunctions.CancelableInvocation`. The function can assign a function to the `oncanceled` property to denote the action to perform when the function is canceled.

If the last function parameter is of type `CustomFunctions.CancelableInvocation`, it will be considered `@cancelable` even if the tag isn't present.

A function can't have both `@cancelable` and `@streaming` tags.

---
### @customfunction
<a id="customfunction"/>

Syntax: @customfunction _id_ _name_

Specify this tag to treat the JavaScript/TypeScript function as an Excel custom function. 

This tag is required to create metadata for the custom function.

There should also be a call to `CustomFunctions.associate("id", functionName);`

The following example shows the simplest way to declare a custom function.

```js
/**
 * Increments a value once a second.
 * @customfunction
 * ...
 */
```

#### id

The `id` is an invariant identifier for the custom function.

* If `id` isn't provided, the JavaScript/TypeScript function name is converted to uppercase and disallowed characters are removed.
* The `id` must be unique for all custom functions.
* The allowed characters are limited to: A-Z, a-z, 0-9, underscores (\_), and period (.).

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

* If name isn't provided, the id is also used as the name.
* Allowed characters: Letters [Unicode Alphabetic character](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), numbers, period (.), and underscore (\_).
* Must start with a letter.
* Maximum length is 128 characters.

In the following example, INC is the `id` of the function and `increment` is the `name`.

```js
/**
 * Increments a value once a second.
 * @customfunction INC INCREMENT
 * ...
 */
```

### description

A description doesn't require any specific tag. Add a description to a custom function by adding a phrase to describe what the function does inside the JSDoc comment. By default, whatever text is untagged in the JSDoc comment section will be the description of the function. The description appears to users in Excel as they are entering the function. In the following example, the phrase "A function that adds two numbers" is the description for the custom function with the id property of `ADD`.

In the following example, ADD is the `id` and `name` of the function and a description is given.

```js
/**
 * A function that adds two numbers.
 * @customfunction ADD
 * ...
 */
```

---
### @helpurl
<a id="helpurl"/>

Syntax: @helpurl _url_

The provided _url_ is displayed in Excel.

In the following example, the `helpurl` is www.contoso.com/weatherhelp.

```js
/**
 * A function which streams the temperature in a town you specify.
 * @customfunction getTemperature
 * @helpurl www.contoso.com/weatherhelp
 * ...
 */
```

---
### @param
<a id="param"/>

#### JavaScript

JavaScript Syntax: @param {type} name _description_

* `{type}` should specify the type info within curly braces. See the [Types](#types) section for more information about the types which may be used. Optional: if not specified, the type `any` will be used.
* `name` specifies which parameter the @param tag applies to. Required.
* `description` provides the description which appears in Excel for the function parameter. Optional.

To denote a custom function parameter as optional:
* Put square brackets around the parameter name. For example: `@param {string} [text] Optional text`.

> [!NOTE]
> The default value for optional parameters is `null`.

The following example shows a ADD function which adds two or three numbers, with the third number as an optional parameter.

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

TypeScript Syntax: @param name _description_

* `name` specifies which parameter the @param tag applies to. Required.
* `description` provides the description which appears in Excel for the function parameter. Optional.

See the [Types](#types) section for more information about the function parameter types which may be used.

To denote a custom function parameter as optional, do one of the following:
* Use an optional parameter. For example: `function f(text?: string)`
* Give the parameter a default value. For example: `function f(text: string = "abc")`

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

---
### @requiresAddress
<a id="requiresAddress"/>

Indicates that the address of the cell where the function is being evaluated should be provided.

The last function parameter must be of type `CustomFunctions.Invocation` or a derived type. When the function is called, the `address` property will contain the address. For an example of a function that uses the `@requiresAddress` tag, see [Addressing cell's context parameter](custom-functions-parameter-options.md#addressing-cells-context-parameter).

---
### @returns
<a id="returns"/>

Syntax: @returns {_type_}

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

---
### @streaming
<a id="streaming"/>

Used to indicate that a custom function is a streaming function. 

The last parameter should be of type `CustomFunctions.StreamingInvocation<ResultType>`.
The function should return `void`.

Streaming functions don't return values directly, but rather should call `setResult(result: ResultType)` using the last parameter.

Exceptions thrown by a streaming function are ignored. `setResult()` may be called with Error to indicate an error result. For an example of a streaming function and more information, see [Make a streaming function](./custom-functions-web-reqs.md#make-a-streaming-function).

Streaming functions can't be marked as [@volatile](#volatile).

---
### @volatile
<a id="volatile"/>

A volatile function is one whose result isn't the same from one moment to the next, even if it takes no arguments or the arguments haven't changed. Excel re-evaluates cells that contain volatile functions, together with all dependents, every time that a calculation is done. For this reason, too much reliance on volatile functions can make recalculation times slow, so use them sparingly.

Streaming functions can't be volatile.

The following function is volatile and uses the `@volatile` tag.

```js
/**
 * Simulates rolling a 6-sided dice.
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

### Matrix type

Use a two-dimensional array type to have the parameter or return value be a matrix of values. For example, the type `number[][]` indicates a matrix of numbers. `string[][]` indicates a matrix of strings.

### Error type

A non-streaming function can indicate an error by returning an Error type.

A streaming function can indicate an error by calling `setResult()` with an Error type.

### Promise

A function can return a Promise, which will provide the value when the promise is resolved. If the promise is rejected, then it is an error.

### Other types

Any other type will be treated as an error.

## Next steps
Learn about [naming conventions for custom functions](custom-functions-naming.md). Alternatively, learn how to [localize your functions](custom-functions-localize.md) which requires you to [write your JSON file by hand](custom-functions-json.md).

## See also

* [Custom functions metadata](custom-functions-json.md)
* [Custom functions best practices](custom-functions-best-practices.md)
* [Create custom functions in Excel](custom-functions-overview.md)
