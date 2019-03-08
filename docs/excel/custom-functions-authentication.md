---
ms.date: 03/06/2019
description: Authenticate users using custom functions in Excel.
title: Authentication for Custom Functions
---

# Authentication

In some scenarios your custom function will need to authenticate the user in order to access protected resources. While custom functions doesn't require a specific method of authentication, you should be aware that custom functions run in a separate runtime from the task pane and other UI elements of your add-in. Because of this, you'll need to pass data back and forth between the two runtimes using the `AsyncStorage` object and the Dialog API.
  
## AsyncStorage object

The custom functions runtime doesn't have a `localStorage` object available on the global window, where you might typically store data. Instead, you should share data between custom functions and task panes by using [OfficeRuntime.AsyncStorage](https://docs.microsoft.com/javascript/api/office-runtime/officeruntime.asyncstorage) to set and get data.

Additionally, there is a benefit to using `AsyncStorage`; it uses a secure sandbox environment so that your data cannot be accessed by other add-ins.

### Suggested usage

When you need to authenticate either from the task pane or a custom function, check `AsyncStorage` to see if the access token was already acquired. If not, use the dialog API to authenticate the user, retrieve the access token, and then store the token in `AsyncStorage` for future use.

## Dialog API

If a token doesn't exist, you should use the Dialog API to ask the user to sign in. After a user enters their credentials, the resulting access token can be stored in `AsyncStorage`.

> [!NOTE]
> The custom functions runtime uses a Dialog object that is slightly different from the Dialog object in the browser engine runtime used by task panes. They're both referred to as the "Dialog API", but use `Officeruntime.Dialog` to authenticate users in the custom functions runtime.

For information on how to use the `OfficeRuntime.Dialog`, see [Custom Functions runtime](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-runtime?view=office-js#displaying-a-dialog-box).

When envisioning the entire authentication process as a whole, it might be helpful to think of the task pane and UI elements of your add-in and the custom functions portions of your add-in as separate entities which can communicate with each other through `AsyncStorage`.

The following diagram outlines this basic process. Note that the dotted line indicates that while they perform separate actions, custom functions and your add-in's task pane are both parts of your add-in as a whole.

1. You issue a custom function call from a cell in an Excel workbook.
2. The custom function uses `Officeruntime.Dialog` to pass your user credentials to a website.
3. This website then returns an access token to the custom function.
4. Your custom function then sets this access token to the `AsyncStorage`.
5. Your add-in's task pane accesses the token from `AsyncStorage`.

![Diagram of custom functions, OfficeRuntime, and task panes working together.](../images/Authdiagram.png "Authentication diagram.")

## Storing the token

The following examples are from the [Using AsyncStorage in custom functions](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) code sample. Refer to this code sample for a complete example of sharing data between custom functions and the task pane.

If the custom function authenticates, then it receives the access token and will need to store it in `AsyncStorage`. The following code sample shows how to call the `AsyncStorage.setItem` method to store a value. The `StoreValue` function is a custom function that for example purposes stores a value from the user. You can modify this to store any token value you need.

```javascript
function StoreValue(key, value) {
  return OfficeRuntime.AsyncStorage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to AsyncStorage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to AsyncStorage. " + error;
  });
}
```

When the task pane needs the access token, it can retrieve the token from `AsyncStorage`. The following code sample shows how to use the `AsyncStorage.getItem` method to retrieve the token.

```javascript
function ReceiveTokenFromCustomFunction() {
   var key = "token";
   var tokenSendStatus = document.getElementById('tokenSendStatus');
   OfficeRuntime.AsyncStorage.getItem(key).then(function (result) {
      tokenSendStatus.value = "Success: Item with key '" + key + "' read from AsyncStorage.";
      document.getElementById('tokenTextBox2').value = result;
   }, function (error) {
      tokenSendStatus.value = "Error: Unable to read item with key '" + key + "' from AsyncStorage. " + error;
   });
}
```

## General guidance

Office Add-ins are web-based and you can use any web authentication technique. There is no particular pattern or method you must follow to implement your own authentication with custom functions. You may wish to consult the documentation about various authentication patterns, starting with [this article about authorizing via external services](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/auth-external-add-ins?view=office-js).  

Avoid using the following locations to store data when developing custom functions:  

- `localStorage`: Custom functions do not have access to the global `window` object and therefore have no access to data     stored in `localStorage`.
- `Office.context.document.settings`:  This location is not secure and information can be extracted by anyone using the     add-in.

## See also

* [Custom functions metadata](custom-functions-json.md)
* [Runtime for Excel custom functions](custom-functions-runtime.md)
* [Custom functions best practices](custom-functions-best-practices.md)
* [Excel custom functions tutorial](excel-tutorial-custom-functions.md)
