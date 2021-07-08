---
ms.date: 05/17/2020
description: Authenticate users using custom functions in Excel which don't use the task pane.
title: Authentication for UI-less custom functions
localization_priority: Normal
---

# Authentication for UI-less custom functions

In some scenarios your custom function that does not use a task pane or other user interface elements (UI-less custom function) will need to authenticate the user in order to access protected resources. Be aware that UI-less custom functions run in a JavaScript-only runtime. Because of this, you'll need to pass data back and forth between the JavaScript-only runtime and the typical browser engine runtime used by most add-ins using the `OfficeRuntime.storage` object and the Dialog API.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

## OfficeRuntime.storage object

The JavaScript-only runtime used by UI-less custom functions doesn't have a `localStorage` object available on the global window, where you typically store data. Instead, you should share data between UI-less custom functions and task panes by using [OfficeRuntime.storage](/javascript/api/office-runtime/officeruntime.storage) to set and get data.

### Suggested usage

When you need to authenticate from a UI-less custom function, check `storage` to see if the access token was already acquired. If not, use the dialog API to authenticate the user, retrieve the access token, and then store the token in `storage` for future use.

## Dialog API

If a token doesn't exist, you should use the Dialog API to ask the user to sign in. After a user enters their credentials, the resulting access token can be stored in `storage`.

> [!NOTE]
> The JavaScript-only runtime uses a Dialog object that is slightly different from the Dialog object in the browser engine runtime used by task panes. They're both referred to as the "Dialog API", but use `OfficeRuntime.Dialog` to authenticate users in the JavaScript-only runtime.

The following diagram outlines this basic process. The dotted line indicates that UI-less custom functions and your add-in's task pane are both part of your add-in as a whole, though they use separate runtimes.

1. You issue a UI-less custom function call from a cell in an Excel workbook.
2. The UI-less custom function uses `Dialog` to pass your user credentials to a website.
3. This website then returns an access token to the UI-less custom function.
4. Your UI-less custom function then sets this access token to the `storage`.
5. Your add-in's task pane accesses the token from `storage`.

![Diagram of custom function using dialog API to get access token, and then share token with task pane through the OfficeRuntime.storage API.](../images/authentication-diagram.png "Authentication diagram.")

## Storing the token

The following examples are from the [Using OfficeRuntime.storage in custom functions](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) code sample. Refer to this code sample for a complete example of sharing data between UI-less custom functions and the task pane.

If the UI-less custom function authenticates, then it receives the access token and will need to store it in `storage`. The following code sample shows how to call the `storage.setItem` method to store a value. The `storeValue` function is a UI-less custom function that for example purposes stores a value from the user. You can modify this to store any token value you need.

```js
/**
 * Stores a key-value pair into OfficeRuntime.storage.
 * @customfunction
 * @param {string} key Key of item to put into storage.
 * @param {*} value Value of item to put into storage.
 */
function storeValue(key, value) {
  return OfficeRuntime.storage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to storage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to storage. " + error;
  });
}
```

When the task pane needs the access token, it can retrieve the token from `storage`. The following code sample shows how to use the `storage.getItem` method to retrieve the token.

```js
/**
 * Read a token from storage.
 * @customfunction GETTOKEN
 */
function receiveTokenFromCustomFunction() {
  var key = "token";
  var tokenSendStatus = document.getElementById('tokenSendStatus');
  OfficeRuntime.storage.getItem(key).then(function (result) {
     tokenSendStatus.value = "Success: Item with key '" + key + "' read from storage.";
     document.getElementById('tokenTextBox2').value = result;
  }, function (error) {
     tokenSendStatus.value = "Error: Unable to read item with key '" + key + "' from storage. " + error;
  });
}
```

## General guidance

Office Add-ins are web-based and you can use any web authentication technique. There is no particular pattern or method you must follow to implement your own authentication with UI-less custom functions. You may wish to consult the documentation about various authentication patterns, starting with [this article about authorizing via external services](../develop/auth-external-add-ins.md).  

Avoid using the following locations to store data when developing custom functions: .

- `localStorage`: UI-less custom functions do not have access to the global `window` object and therefore have no access to data stored in `localStorage`.
- `Office.context.document.settings`:  This location is not secure and information can be extracted by anyone using the add-in.

## Dialog box API example

In the following code sample, the function `getTokenViaDialog` uses the `Dialog` API's `displayWebDialogOptions` function to display a dialog box. This sample is provided to show the capabilities of the `Dialog` object, not demonstrate how to authenticate.

```JavaScript
/**
 * Function retrieves a cached token or opens a dialog box if there is no saved token. Note that this is not a sufficient example of authentication but is intended to show the capabilities of the Dialog object.
 * @param {string} url URL for a stored token.
 */
function getTokenViaDialog(url) {
  return new Promise (function (resolve, reject) {
    if (_dialogOpen) {
      // Can only have one dialog box open at once. Wait for previous dialog box's token.
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
      OfficeRuntime.displayWebDialog(url, {
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
Learn how to [debug UI-less custom functions](custom-functions-debugging.md).

## See also

* [Runtime for UI-less Excel custom functions](custom-functions-runtime.md)
* [Excel custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md)