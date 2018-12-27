---
ms.date: 12/26/2018
description: Learn how to authenticate users via custom functions in Excel.
title: Authentication for Custom Functions
---

# Authentication

You may wish to verify that a user is authenticated before allowing them access to your custom functions. Implementing authentication for custom functions does not differ significantly in process from other Office add-ins. For specific details on the process of authenticating with single sign-on, see this article [Enable single sign-on for Office Add-ins (preview)](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/sso-in-office-add-ins).  
  
However, there are some specific APIs used by custom functions which you should use when following the recommended authentication process.  
  
## AsyncStorage

The `AsyncStorage` is common and accessible to both custom functions and UI elements of your add-in such as the task pane. Because of this commonality, you can use it to store information which needs to pass back and forth between these parts of your add in. For example, if a user enters their credentials through a UI element, resulting access and refresh tokens can be captured and stored in `AsyncStorage` which also makes this information available to custom functions.

Additionally,  `AsyncStorage` offers a sandboxed environment on a user's device and cannot be accessed by other add-ins.  
  
There are some locations which should not be used to store data if you are using custom functions:  

    - `localStorage`: Custom functions do not have access to the global `window` object and therefore have no access to data stored in `localStorage`.

    - `Office.context.document.settings`:  `Office.context.document.settings` is not secure and information can be extracted by anyone using the add-in.

## Dialog API

If your function checks `AsyncStorage` and does not find an access token in the process of authenticating, you should use the DisplayWebDialog API to prompt the user to enter their credentials.  
  
As mentioned previously, when a user enters their credentials via the dialog box, tokens should be stored using `AsyncStorage`.  
  
The following code sample shows how you can use the `displayWebDialog` to display a dialog box   for that purpose.

```js
var jwt = require('jsonwebtoken')
var secret = process.env.secret // sample assumes you have already set a secret here

// getStock function, which is exposed to the user in Excel
function getStock (ticker) {
    return new Promise(function (resolve, reject) {
      // Get a token
      getToken("https://www.contoso.com/auth")
      .then(function (token) {

        // Use token to get stock price
        fetch("https://www.contoso.com/?token=token&ticker= + ticker")
        .then(function (result) {

          // Return stock price to cell
          resolve(result);
        });
      })
      .catch(function (error) {
        reject(error);
      });
    });

    //Helper method within getStock to get the token
    function getToken(url) {
        return new Promise(function (resolve,reject) {
        if(_cachedToken) {
            resolve(_cachedToken);
        } else { 
            getTokenViaDialog(url)
            .then(function (result) {
            resolve(result);
            })
            .catch(function (result) {
            reject(result);
            });
        }
        });
    };

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
                // Open a dialog
                _dialogOpen = true;
               OfficeRuntime.displayWebDialogOptions(url, {
                  height: '50%',
                  width: '50%',
                  onMessage: function (message, dialog) {
                    let parsedToken = JSON.stringify(message);
                    var decodedToken = jwt.verify(parsedToken, secret);
                    dialog.close(); 
                    resolve(decodedToken);
                  },
                  onRuntimeError: function(error, dialog) {
                    reject(error);
                    dialog.close();
                  },
                }).catch(function (e) {
                  reject(e);
                });
              }
            });
    }
}

```
