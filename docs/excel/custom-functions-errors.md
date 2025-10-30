---
title: Handle and return errors from your custom function
description: 'Return meaningful Excel errors (like #VALUE! and #N/A) from custom functions and map exceptions to user-friendly messages.'
ms.date: 10/22/2025
ms.localizationpriority: medium
---

# Handle and return errors from your custom function

When a custom function receives invalid input, can't access a resource, or fails to compute a result, return the most specific Excel error you can. Validate parameters early to fail promptly and use `try...catch` blocks to turn low-level exceptions into clear Excel errors.

## Detect and throw an error

The following example validates a U.S. ZIP Code with a regular expression before continuing. If the format is invalid, it throws a `#VALUE!` error.

```typescript
/**
* Gets a city name for the given U.S. ZIP Code.
* @customfunction
* @param {string} zipCode
* @returns The city of the ZIP Code.
*/
function getCity(zipCode: string): string {
  let isValidZip = /(^\d{5}$)|(^\d{5}-\d{4}$)/.test(zipCode);
  if (isValidZip) return cityLookup(zipCode);
  let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, "Please provide a valid U.S. zip code.");
  throw error;
}
```

## The CustomFunctions.Error object

The [CustomFunctions.Error](/javascript/api/custom-functions-runtime/customfunctions.error) object returns an error to the cell. Specify which error by choosing an `ErrorCode` value from the following list.

|ErrorCode enum value  |Excel cell value  |Description  |
|---------------|---------|---------|
|`divisionByZero` | `#DIV/0`  | The function is attempting to divide by zero. |
|`invalidName`    | `#NAME?`  | There is a typo in the function name. Note that this error is supported as a custom function input error, but not as a custom function output error. |
|`invalidNumber`  | `#NUM!`   | There is a problem with a number in the formula. |
|`invalidReference` | `#REF!` | The function refers to an invalid cell. Note that this error is supported as a custom function input error, but not as a custom function output error.|
|`invalidValue`   | `#VALUE!` | A value in the formula is of the wrong type. |
|`notAvailable`   | `#N/A`    | The function or service isn't available. |
|`nullReference`  | `#NULL!`  | The ranges in the formula don't intersect. |

The following code sample shows how to create and return an error for an invalid number (`#NUM!`).

```typescript
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidNumber);
throw error;
```

The `#VALUE!` and `#N/A` errors also support custom error messages. Custom error messages are displayed in the error indicator menu, which is accessed by hovering over the error flag on each cell with an error. The following example shows how to return a custom error message with the `#VALUE!` error.

```typescript
// You can only return a custom error message with the #VALUE! and #N/A errors.
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, "The parameter can only contain lowercase characters.");
throw error;
```

### Handle errors when working with dynamic arrays

Custom functions can return dynamic arrays that include errors. For example, a custom function could output the array `[1],[#NUM!],[3]`. The following code sample shows how to pass three parameters into a custom function, replace one parameter with a `#NUM!` error, and then return a two-dimensional array with the results for each input.

```js
/**
* Returns the #NUM! error as part of a 2-dimensional array.
* @customfunction
* @param {number} first First parameter.
* @param {number} second Second parameter.
* @param {number} third Third parameter.
* @returns {number[][]} Three results, as a 2-dimensional array.
*/
function returnInvalidNumberError(first, second, third) {
  // Use the `CustomFunctions.Error` object to retrieve an invalid number error.
  const error = new CustomFunctions.Error(
    CustomFunctions.ErrorCode.invalidNumber, // Corresponds to the #NUM! error in the Excel UI.
  );

  // Enter logic that processes the first, second, and third input parameters.
  // Imagine that the second calculation results in an invalid number error. 
  const firstResult = first;
  const secondResult =  error;
  const thirdResult = third;

  // Return the results of the first and third parameter calculations and a #NUM! error in place of the second result. 
  return [[firstResult], [secondResult], [thirdResult]];
}
```

### Errors as custom function inputs

A custom function can evaluate even if the input range contains an error. For example, a custom function can take the range **A2:A7** as an input, even if **A6:A7** contains an error.

To process inputs that contain errors, a custom function must have the JSON metadata property `allowErrorForDataTypeAny` set to `true`. See [Manually create JSON metadata for custom functions](custom-functions-json.md#metadata-reference) for more information.

> [!IMPORTANT]
> The `allowErrorForDataTypeAny` property can only be used with [manually created JSON metadata](custom-functions-json.md). This property doesn't work with the autogenerated JSON metadata process.

## Use `try...catch` blocks

Use [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) blocks to catch potential errors and return meaningful error messages to your users. By default, Excel returns `#VALUE!` for unhandled errors or exceptions.

In the following code sample, the custom function uses fetch to call a REST service. If the call fails, such as when the REST service returns an error or the network is unavailable, the custom function returns `#N/A` to show that the web call failed.

```typescript
/**
 * Gets a comment from the hypothetical contoso.com/comments API.
 * @customfunction
 * @param {number} commentID ID of a comment.
 */
function getComment(commentID) {
  let url = "https://www.contoso.com/comments/" + commentID;
  return fetch(url)
    .then(function (data) {
      return data.json();
    })
    .then(function (json) {
      return json.body;
    })
    .catch(function (error) {
      throw new CustomFunctions.Error(CustomFunctions.ErrorCode.notAvailable);
    })
}
```

## Next steps

Learn how to [troubleshoot problems with your custom functions](custom-functions-troubleshooting.md).

## See also

* [Custom functions debugging](custom-functions-debugging.md)
* [Custom functions requirement sets](/javascript/api/requirement-sets/excel/custom-functions-requirement-sets)
* [Create custom functions in Excel](custom-functions-overview.md)
