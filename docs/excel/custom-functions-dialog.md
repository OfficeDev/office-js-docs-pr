---
ms.date: 05/06/2019
description: Create a dialog box through custom functions in Excel using JavaScript.
title: Display a dialog box from a custom function
localization_priority: Priority
---
# Display a dialog box from a custom function

If your custom function needs to interact with the user, you can create a dialog box using the [`Office.Dialog` object](/javascript/api/office-runtime/officeruntime.dialog?view=office-js). A common scenario for using the dialog box is to authenticate a user so that your custom function can access a web service. For more details about authentication with custom functions, see [Custom functions authentication](./custom-functions-authentication.md).

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

>[!NOTE]
> The `Office.Dialog` object is part of the custom functions runtime. Task panes don't use the `Dialog` object. To create a dialog box from a task pane, see [Dialog API](/office/dev/add-ins/develop/dialog-api-in-office-add-ins).

## dialog box API example

In the following code sample, the function `getTokenViaDialog` uses the `Dialog` API’s `displayWebDialogOptions` function to display a dialog box.

```js
/**
 * Function retrieves a cached token or opens a dialog box if there is no saved token. Note that this is not a sufficient example of authentication but is intended to show the capabilities of the Dialog object.
 * @param {string} url URL for a stored token.
 */
function getTokenViaDialog(url) {
  return new Promise (function (resolve, reject) {
    if (_dialogOpen) {
      // Can only have one dialog box open at once, wait for previous dialog box's token
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

## Next steps
Learn how to [make your custom functions compatible with XLL user-defined functions](make-custom-functions-compatible-with-xll-udf.md).

## See also

* [Custom functions authentication](custom-functions-authentication.md)
* [Receive and handle data with custom functions](custom-functions-web-reqs.md)
* [Create custom functions in Excel](custom-functions-overview.md)
