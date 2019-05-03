---
ms.date: 05/03/2019
description: Handle errors in your Excel custom functions.
title: Error handling for custom functions in Excel (preview)
localization_priority: Priority
---

# Error handling within custom functions (preview)

When you build an add-in that defines custom functions, be sure to include error handling logic to account for runtime errors. Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

In the following code sample, `.catch` will handle any errors that occur previously in the code.

```js
/**
 * Gets a comment from the hypothetical contoso.com/comments API.
 * @customfunction
 * @param {number} commentID ID of a comment.
 */
function getComment(commentID) {
  let url = "https://www.contoso.com/comments/" + x;

  return fetch(url)
    .then(function (data) {
      return data.json();
    })
    .then(function (json) {
      return json.body;
    })
    .catch(function (error) {
      throw error;
    })
}
```

## Next steps
Learn how to [troubleshoot problems with your custom functions](custom-functions-troubleshooting.md).

## See also

* [Custom functions debugging](custom-functions-debugging.md)
* [Custom functions changelog](custom-functions-changelog.md)
* [Custom functions requirements](custom-functions-requirements.md)
* [Create custom functions in Excel](custom-functions-overview.md)
