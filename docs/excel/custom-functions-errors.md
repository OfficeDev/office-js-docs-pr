---
ms.date: 10/31/2019
description: 'Handle and return errors like #NULL! from your custom function'
title: Handle and return errors from your custom function (preview)
localization_priority: Priority
---

# Handle and return errors from your custom function (preview)

> [!NOTE]
> The features described in this article are currently in preview and subject to change. They are not currently supported for use in production environments. You will need to [Office Insider](https://insider.office.com/en-us/join) to try the preview features.  A good way to try out preview features is by using an Office 365 subscription. If you don't already have an Office 365 subscription, you can get one by joining the [Office 365 Developer Program](https://developer.microsoft.com/office/dev-program).

If something goes wrong while your custom function runs, you will need to return an error to inform the user. If you have specific parameter requirements, such as only positive numbers, you will need to test the parameters and throw an error if they are not correct. You can also use a `try`-`catch` block to catch any errors that occur while your custom function runs.

## Detect and throw an error

Let’s look at a case where you need to ensure that a zip code parameter is in the correct format for the custom function to work. The following custom function uses a regular expression to check the zip code. If it is correct, then it will look up the city (in another function) and return the value. If it is not correct, it returns a `#VALUE!` error to the cell.

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

The `CustomFunctions.Error` object is used to return an error back to the cell. When you create the object, specify which error you want to use by using one of the following `ErrorCode` enum values.


|ErrorCode enum value  |Excel cell value  |Meaning  |
|---------------|---------|---------|
|`invalidValue`   | `#VALUE!` | A value used in the formula is the wrong type. |
|`notAvailable`   | `#N/A`    | The function or service is not available. |
|`divisionByZero` | `#DIV/0`  | Be aware that JavaScript allows division by zero so you need to write an error handler carefully to detect this condition. |
|`invalidNumber`  | `#NUM!`   | There is a problem with the number used in the formula |
|`nullReference`  | `#NULL!`  | The ranges in the formula do not intersect. |

The following code sample shows how to create and return an error for an invalid number (`#NUM!`).

```typescript
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidNumber);
throw error;
```

When you return a `#VALUE!` error you can also include a custom message that will be shown in a popup when the user hovers over the cell. The following example shows how to return a custom error message.

```typescript
// You can only return a custom error message with the #VALUE! error
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, “The parameter can only contain lowercase characters.”);
throw error;
```

## Use try-catch blocks

In general, you should use `try`-`catch` blocks in your custom function to catch any potential errors that occur. If you do not handle exceptions in your code, they will be returned to Excel. By default, Excel returns `#VALUE!` for an unhandled exception.

In the following code sample, the custom function makes a fetch call to a REST service. It's possible that the call will fail, for example, if the REST service returns an error or the network goes down. If this happens, the custom function will return `#N/A` to indicate the web call failed.


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
* [Custom functions requirements](custom-functions-requirement-sets.md)
* [Create custom functions in Excel](custom-functions-overview.md)
