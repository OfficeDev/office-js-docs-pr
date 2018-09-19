---
ms.date: 09/18/2018
description: Learn best practices and recommended patterns for Excel custom functions.
title: Custom Functions' best practices
---

# Custom Functions' best practices

This article describes some recommended patterns and solutions to common use cases with Excel custom functions.

## Authentication

Add-ins that contain web views (such as a task pane), but do not include any custom functions accomplish authentication via `Office.context.displayDialogAsync`. Because custom functions use a different runtime, they do not have access to the `Office.context` methods. Instead, they use methods available on the `OfficeRuntime` object, such as `OfficeRuntime.DisplayWebDialog()` and `OfficeRuntime.AsyncStorage.setItem(key, value)`.

The following code sample shows how to store a token in AsyncStorage and display a dialog box which can indicate whether or not the user is authenticated.  

```js
// Get auth token before calling my service, a hypothetical API, which will deliver a stock price based on stock ticker string, such as "MSFT"
async function getStock(ticker) {
    const token = await getToken();
    let data = await (await fetch(https://myservice.com/?token=token&ticker= + ticker).json());
    return data.price;
}

async function getToken() {
    if (_cachedToken) {
        return _cachedToken;
    } else {
        return await getTokenViaDialog_AsPromise();
    }
}
  
// Function to display dialog window
function getTokenViaDialog_AsPromise() {
    return new Promise ((resolve, reject) => {
        displayWebDialog("https://www.auth.com/", {
           height: ’50’,
           width: ’50%’,
           hideTitle: true,
           onMessage: (message, dialog) => {
               let json = JSON.parse(message);
                if (json.type === "token_succeeded") {
                        resolve(json.value);
                        dialog.closeDialog();
                        return;
                    }
            // Otherwise, handle other messages.
            },
          onClose: () => reject("User closed dialog")
        }).catch(e => reject(e));
    });
}
```

Looking for information about authentication not specific to Excel custom functions? See the article on [authorizing external services for add-ins which do not use the new JavaScript runtime](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/auth-external-add-ins).

## Batching of API requests

Instead of executing individual web requests, it is possible to batch requests using custom functions.

The below gives an outline of a batching pattern you can modify for your custom function:

```js
// Hypothetical public async API
function sumAsync() {
    // Push an entry into the next batch
    var batchEntry = {
      "arguments": arguments,
      "resolve": undefined,
      "reject": undefined
    };
    var promise = new Promise((resolve, reject) => {
      batchEntry.resolve = resolve;
      batchEntry.reject = reject;
    });
    _batch.push(batchEntry);
  
    // If a batch hasn't been scheduled yet, schedule it after a certain timeout, e.g. 2 seconds
    if (!_isBatchScheduled) {
      setTimeout(_processBatch, 2000);
      _isBatchScheduled = true;
    }
  
    // Return the promise
    return promise;
  }
  
  // Current batch
  var _batch = [];
  var _isBatchScheduled = false;
  
  // Internal batch processor
  function _processBatch() {
    // If there is anything batched...
    if (_batch.length > 0) {
      for (var i = 0; i < _batch.length; i++) {

        // Sum up the arguments
        var sum = 0;
        for (var j = 0; j < _batch[i].arguments.length; j++) {
          sum += _batch[i].arguments[j];
        }

        // In this simple example we always resolve, but you should also add error handling for reject
        _batch[i].resolve(sum);
      }

      // The current batch has been processed
      _batch = [];
      _isBatchScheduled = false;
    }
  }
  
  
  // Sample usage
  console.clear();
  sumAsync(1, 2).then((value) => { console.log("1 + 2 = " + value); });
  sumAsync(2, 4, 6).then((value) => { console.log("2 + 4 + 6 = " + value); });
  sumAsync(3, 5, 7, 9).then((value) => { console.log("3 + 5 + 7 + 9 = " + value); });
```

## Error handling

Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](./excel-add-ins-error-handling.md). Generally, you will use `.catch` to handle errors. The code below gives an example of `.catch`.

```js
function getComment(x) {
    var url = "https://jsonplaceholder.typicode.com/comments/" + x; //this delivers a section of lorem ipsum from the jsonplaceholder API
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

## Error logging

You can use runtime logging to debug your custom function's XML manifest file or to look for errors in your custom functions in real time via console.log statements. Runtime logging is only available for Office 2016 desktop currently.

For full instructions on how to use runtime logging, [read this article](../testing/troubleshoot-manifest.md).

## Debugging

At present, the best method for debugging Excel custom functions is to use [Excel Online](https://www.office.com/launch/excel) and use the F12 debugging tool native to your browser. Additional debugging tools for custom functions may be available in the future.

## Mapping names

Custom functions are typically declared entirely in uppercase letters, although you can change this by using the  `CustomFunctionsMappings` object. The key-value pairs you specify in `CustomFunctionsMappings` correspond to the function name you call in Excel (such as `=ADD42`) and the new alternate name you would like to use for this function in Excel. Use of `CustomFunctionsMapping` is not required, but can be helpful if you are using an uglifier, webpack, or import syntax - all of which have difficulty with the uppercase letters in these functions.
  
You can declare individual functions, as shown below:  

```js
function ADD42(num) {
    return num + 42;  
}  
  
CustomFunctionsMappings = {
    "plusFortyTwo" : ADD42 //effectively renames the add-in when invoked in Excel, so you will now call =plusFortyTwo()
}
```

However, you can declare multiple mappings at the same time, as shown in the example below.  

```js
//assume that COUNTDOGS and COUNTCATS exist
  
CustomFunctionsMappings = {
    "countdogs" : COUNTDOGS,  
    "meow" : COUNTCATS
}
 ```