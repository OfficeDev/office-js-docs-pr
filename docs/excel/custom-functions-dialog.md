---
ms.date: 05/01/2019
description: Create dialogs through custom functions in Excel using JavaScript.
title: Custom functions dialogs (preview)
localization_priority: Priority
---
# Display a dialog in custom functions

If your custom function needs to interact with the user, you can create a dialog using the [`Office.Dialog` object](/javascript/api/office-runtime/officeruntime.dialog?view=office-js). A common scenario for using the dialog is to authenticate a user so that your custom function can access a web service. For more details about authentication with custom functions, see [Custom functions authentication](./custom-functions-authentication.md).

>[!NOTE]
> The `Office.Dialog` object is part of the custom functions runtime. Task panes don't use the `Dialog` object. To create a dialog from a task pane, see [Dialog API](/office/dev/add-ins/develop/dialog-api-in-office-add-ins).

## Dialog API example

In the following code sample, the function `getTokenViaDialog` uses the `Dialog` API’s `displayWebDialogOptions` function to display a dialog.

```js
/**
 * Function retrieves a cached token or opens a dialog if there is no saved token. Note that this is not a sufficient example of authentication but is intended to show the capabilities of the Dialog object.
 * @param {string} url URL for a stored token.
 */
function getTokenViaDialog(url) {
  return new Promise (function (resolve, reject) {
    if (_dialogOpen) {
      // Can only have one dialog open at once, wait for previous dialog's token
      let timeout = 5;
      let count = 0;
      var intervalId = setInterval(function () {
        count++;
        if(_cachedToken) {
          resolve(_cachedToken);
          clearInterval(intervalId);
        }
        if(count >= timeout) {
          reject("Timeout while waiting for token");
          clearInterval(intervalId);
        }
      }, 1000);
    } else {
      _dialogOpen = true;
      Office.displayWebDialogOptions(url, {
        height: '50%',
        width: '50%',
        onMessage: function (message, dialog) {
          _cachedToken = message;
          resolve(message);
          dialog.close();
          return;
        },
        onRuntimeError: function(error, dialog) {
          reject(error);
        },
      }).catch(function (e) {
        reject(e);
      });
    }
  });
}
```

## See also

* [Custom functions metadata](custom-functions-json.md)
* [Runtime for Excel custom functions](custom-functions-runtime.md)
* [Custom functions best practices](custom-functions-best-practices.md)
* [Custom functions changelog](custom-functions-changelog.md)
* [Excel custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md)
