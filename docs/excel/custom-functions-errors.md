---
title: Handle and return errors from your custom function
description: 'Handle and return errors like #NULL! from your custom function.'
ms.date: 08/12/2021
ms.localizationpriority: medium
---

# Handle and return errors from your custom function

If something goes wrong while your custom function runs, return an error to inform the user. If you have specific parameter requirements, such as only positive numbers, test the parameters and throw an error if they aren't correct. You can also use a [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) block to catch any errors that occur while your custom function runs.

## Detect and throw an error

Let's look at a case where you need to ensure that a zip code parameter is in the correct format for the custom function to work. The following custom function uses a regular expression to check the zip code. If the zip code format is correct, then it will look up the city using another function and return the value. If the format isn't valid, the function returns a `#VALUE!` error to the cell.

```typescript
/**
* Gets a city name for the given U.S. zip code.
* @customfunction
* @param {string} zipCode
* @returns The city of the zip code.
*/
function getCity(zipCode: string): string {
  let isValidZip = /(^\d{5}$)|(^\d{5}-\d{4}$)/.test(zipCode);
  if (isValidZip) return cityLookup(zipCode);
  let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, "Please provide a valid U.S. zip code.");
  throw error;
}
```

## The CustomFunctions.Error object

The [CustomFunctions.Error](/javascript/api/custom-functions-runtime/customfunctions.error) object is used to return an error back to the cell. When you create the object, specify which error you want to use by choosing one of the following `ErrorCode` enum values.

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

In addition to returning a single error, a custom function can output a dynamic array that includes an error. For example, a custom function could output the array `[1],[#NUM!],[3]`. The following code sample shows how to input three parameters into a custom function, replace one of the input parameters with a `#NUM!` error, and then return a 2-dimensional array with the results of processing each input parameter.

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

In general, use [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) blocks in your custom function to catch any potential errors that occur. If you don't handle exceptions in your code, they will be returned to Excel. By default, Excel returns `#VALUE!` for unhandled errors or exceptions.

In the following code sample, the custom function makes a fetch call to a REST service. It's possible that the call will fail, for example, if the REST service returns an error or the network goes down. If this happens, the custom function will return `#N/A` to indicate that the web call failed.

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
