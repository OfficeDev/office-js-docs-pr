---
ms.date: 02/08/2019
description: Handle errors in your Excel custom functions.
title: Error handling for custom functions in Excel (preview)
localization_priority: Priority
---
# Error handling within custom functions

When you build an add-in that defines custom functions, be sure to include error handling logic to account for runtime errors. Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

In the following code sample, `.catch` will handle any errors that occur previously in the code.

```js
function getComment(commentID) {
  let url = "https://www.contoso.com/comments/" + x;

  return fetch(url)
    .then(function (data) {
      return data.json();
    })
    .then((json) => {
      return json.body;
    })
    .catch(function (error) {
      throw error;
    })
}
```

## See also

* [Excel custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md)
* [Custom functions metadata](custom-functions-json.md)
* [Runtime for Excel custom functions](custom-functions-runtime.md)
* [Custom functions best practices](custom-functions-best-practices.md)
* [Custom functions changelog](custom-functions-changelog.md)
