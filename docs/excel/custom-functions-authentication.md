---
ms.date: 05/03/2019
description: Authenticate users using custom functions in Excel.
title: Authentication for custom functions
localization_priority: Priority
---

# Authentication for custom functions

In some scenarios your custom function will need to authenticate the user in order to access protected resources. While custom functions don't require a specific method of authentication, you should be aware that custom functions run in a separate runtime from the task pane and other UI elements of your add-in. Because of this, you'll need to pass data back and forth between the two runtimes using the `OfficeRuntime.storage` object and the Dialog API.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## OfficeRuntime.storage object

The custom functions runtime doesn't have a `localStorage` object available on the global window, where you might typically store data. Instead, you should share data between custom functions and task panes by using [OfficeRuntime.storage](/javascript/api/office-runtime/officeruntime.asyncstorage) to set and get data.

Additionally, there is a benefit to using the `storage` object; it uses a secure sandbox environment so that your data cannot be accessed by other add-ins.

### Suggested usage

When you need to authenticate either from the task pane or a custom function, check `storage` to see if the access token was already acquired. If not, use the dialog API to authenticate the user, retrieve the access token, and then store the token in `storage` for future use.

## Dialog API

If a token doesn't exist, you should use the Dialog API to ask the user to sign in. After a user enters their credentials, the resulting access token can be stored in `storage`.

> [!NOTE]
> The custom functions runtime uses a Dialog object that is slightly different from the Dialog object in the browser engine runtime used by task panes. They're both referred to as the "Dialog API", but use `OfficeRuntime.Dialog` to authenticate users in the custom functions runtime.

For information on how to use the `Dialog` object, see [Custom Functions dialog](/office/dev/add-ins/excel/custom-functions-dialog).

When envisioning the entire authentication process as a whole, it might be helpful to think of the task pane and UI elements of your add-in and the custom functions part of your add-in as separate entities which can communicate with each other through `OfficeRuntime.storage`.

The following diagram outlines this basic process. Note that the dotted line indicates that while they perform separate actions, custom functions and your add-in's task pane are both part of your add-in as a whole.

1. You issue a custom function call from a cell in an Excel workbook.
2. The custom function uses `Dialog` to pass your user credentials to a website.
3. This website then returns an access token to the custom function.
4. Your custom function then sets this access token to the `storage`.
5. Your add-in's task pane accesses the token from `storage`.

![Diagram of custom function using dialog API to get access token, and then share token with task pane through the OfficeRuntime.storage API.](../images/authentication-diagram.png "Authentication diagram.")

## Storing the token

The following examples are from the [Using OfficeRuntime.storage in custom functions](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) code sample. Refer to this code sample for a complete example of sharing data between custom functions and the task pane.

If the custom function authenticates, then it receives the access token and will need to store it in `storage`. The following code sample shows how to call the `storage.setItem` method to store a value. The `storeValue` function is a custom function that for example purposes stores a value from the user. You can modify this to store any token value you need.

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

CustomFunctions.associate("STOREVALUE", storeValue);
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
CustomFunctions.associate("GETTOKEN", receiveTokenFromCustomFunction);

```

## General guidance

Office Add-ins are web-based and you can use any web authentication technique. There is no particular pattern or method you must follow to implement your own authentication with custom functions. You may wish to consult the documentation about various authentication patterns, starting with [this article about authorizing via external services](/office/dev/add-ins/develop/auth-external-add-ins?view=office-js).  

Avoid using the following locations to store data when developing custom functions:  

- `localStorage`: Custom functions do not have access to the global `window` object and therefore have no access to data stored in `localStorage`.
- `Office.context.document.settings`:  This location is not secure and information can be extracted by anyone using the add-in.

## Next steps
Learn about the [dialog API for custom functions](custom-functions-dialog.md).

## See also

* [Custom functions architecture](custom-functions-architecture.md)
* [Receive and handle data with custom functions](custom-functions-web-reqs.md)
* [Runtime for Excel custom functions](custom-functions-runtime.md)
* [Excel custom functions tutorial](excel-tutorial-custom-functions.md)
