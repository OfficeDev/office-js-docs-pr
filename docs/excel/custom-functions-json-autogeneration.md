---
ms.date: 05/03/2019
description: Use JSDOC tags to dynamically create your custom functions JSON metadata.
title: Autogenerate JSON metadata for custom functions
localization_priority: Priority
---

# Autogenerate JSON metadata for custom functions

When an Excel custom function is written in JavaScript or TypeScript, JSDoc tags are used to provide extra information about the custom function. The JSDoc tags are then used at build time to create the [JSON metadata file](custom-functions-json.md). Using JSDoc tags saves you from the effort of manually editing the JSON metadata file.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Add the `@customfunction` tag in the code comments for a JavaScript or TypeScript function to mark it as a custom function.

The function parameter types may be provided using the [@param](#param) tag in JavaScript, or from the [Function type](https://www.typescriptlang.org/docs/handbook/functions.html) in TypeScript. For more information, see the [@param](#param) tag and [Types](#types) section.

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

If the last function parameter is of type `CustomFunctions.CancelableInvocation`, it will be considered `@cancelable` even if the tag is not present.

A function cannot have both `@cancelable` and `@streaming` tags.

---
### @customfunction
<a id="customfunction"/>

Syntax: @customfunction _id_ _name_

Specify this tag to treat the JavaScript/TypeScript function as an Excel custom function.

This tag is required to create metadata for the custom function.

There should also be a call to `CustomFunctions.associate("id", functionName);`

#### id

The id is used as the invariant identifier for the custom function stored in the document. It should not change.

* If id is not provided, the JavaScript/TypeScript function name is converted to uppercase, disallowed characters are removed.
* The id must be unique for all custom functions.
* The characters allowed are limited to: A-Z, a-z, 0-9, and period (.).

#### name

Provides the display name for the custom function.

* If name is not provided, the id is also used as the name.
* Allowed characters: Letters [Unicode Alphabetic character](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), numbers, period (.), and underscore (\_).
* Must start with a letter.
* Maximum length is 128 characters.

---
### @helpurl
<a id="helpurl"/>

Syntax: @helpurl _url_

The provided _url_ is displayed in Excel.

---
### @param
<a id="param"/>

#### JavaScript

JavaScript Syntax: @param {type} name _description_

* `{type}` should specify the type info within curly braces. See the [Types](##types) for more information about the types which may be used. Optional: if not specified, the type `any` will be used.
* `name` specifies which parameter the @param tag applies to. Required.
* `description` provides the description which appears in Excel for the function parameter. Optional.

To denote a custom function parameter as optional:
* Put square brackets around the parameter name. For example: `@param {string} [text] Optional text`.

> [!NOTE]
> The default value for optional parameters is `null`.

#### TypeScript

TypeScript Syntax: @param name _description_

* `name` specifies which parameter the @param tag applies to. Required.
* `description` provides the description which appears in Excel for the function parameter. Optional.

See the [Types](##types) for more information about the function parameter types which may be used.

To denote a custom function parameter as optional, do one of the following:
* Use an optional parameter. For example: `function f(text?: string)`
* Give the parameter a default value. For example: `function f(text: string = "abc")`

For detailed description of the @param see: [JSDoc](https://usejsdoc.org/tags-param.html)

> [!NOTE]
> The default value for optional parameters is `null`.

---
### @requiresAddress
<a id="requiresAddress"/>

Indicates that the address of the cell where the function is being evaluated should be provided. 

The last function parameter must be of type `CustomFunctions.Invocation` or a derived type. When the function is called, the `address` property will contain the address.

---
### @returns
<a id="returns"/>

Syntax: @returns {_type_}

Provides the type for the return value.

If `{type}` is omitted, the TypeScript type info will be used. If there is no type info, the type will be `any`.

---
### @streaming
<a id="streaming"/>

Used to indicate that a custom function is a streaming function. 

The last parameter should be of type `CustomFunctions.StreamingInvocation<ResultType>`.
The function should return `void`.

Streaming functions do not return values directly, but rather should call `setResult(result: ResultType)` using the last parameter.

Exceptions thrown by a streaming function are ignored. `setResult()` may be called with Error to indicate an error result.

Streaming functions cannot be marked as [@volatile](#volatile).

---
### @volatile
<a id="volatile"/>

A volatile function is one whose result cannot be assumed to be the same from one moment to the next even if it takes no arguments or the arguments have not changed. Excel re-evaluates cells that contain volatile functions, together with all dependents, every time that a calculation is done. For this reason, too much reliance on volatile functions can make recalculation times slow, so use them sparingly.

Streaming functions cannot be volatile.

---

## Types

By specifying a parameter type, Excel will convert values into that type before calling the function. If the type is `any`, no conversion will be performed.

### Value types

A single value may be represented using one of the following types: `boolean`, `number`, `string`.

### Matrix type

Use a two-dimensional array type to have the parameter or return value be a matrix of values. For example, the type `number[][]` indicates a matrix of numbers. `string[][]` indicates a matrix of strings. 

### Error type

A non-streaming function can indicate an error by returning an Error type.

A streaming function can indicate an error by calling setResult() with an Error type.

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
