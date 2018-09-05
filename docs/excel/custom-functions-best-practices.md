---
ms.date: 09/05/2018
description: Learn best practices and recommended patterns for Excel custom functions.
title: Custom Functions' best practices
---

# Custom Functions' best practices
This article covers some recommended patterns and solutions to common use cases with Excel custom functions.

## Authentication
Regular add-ins use `Office.context.displayDialogAsync` and `Office.context.asyncStorage` to accomplish authentication tasks. Custom functions differs from this approach. Because custom functions use a different runtime, they do not have access to the Office.context methods. Instead, they use methods available on the `OfficeRuntime` object, such as `OfficeRuntime.DisplayWebDialog()` and `OfficeRuntime.AsyncStorage.setItem(key, value)`.

The following code sample shows how to store a token in AsyncStorage and display a dialog box which can indicate whether or not the user is authenticated.  

```ts
// Get auth token before calling my service, a hypothetical API which will deliver a stock price based on stock ticker string, such as "MSFT"
async function getStock(ticker: string) {
    const token = await getToken();
    let data = await (await fetch(https://myservice.com/?token=token&ticker= + ticker).json());
    return data.price;
}

async function getToken(): Promise<string> {
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
            onClose: () => reject("User closed dialog"),
            onRuntimeErrors: (e) => reject(e)  
        }).catch(e => reject(e));
    });
}
```

## Batching of API requests
Instead of executing individual web requests, it is possible to batch requests using custom functions.

The below gives an outline of a batching pattern you can modify for your custom function:

```js
// Current batch
var _argumentsBatch = [];
var _isBatchScheduled = false;

// Individual requests are batched into an array
function batchRequests(a, b, c) {
  // Push the arguments to a batch
  _argumentsBatch.push(arguments);

  // If a batch hasn't been scheduled, schedule it after one second
  if (!_isBatchScheduled) {
    setTimeout(_processBatch, 1000); //milliseconds
    _isBatchScheduled = true;
  }
}

// Internal batch processor
function _processBatch() {
  // If anything needs to be batched, begin batching.
  if (_argumentsBatch.length > 0) {
    for (var i = 0; i < _argumentsBatch.length; i++) {
      // Process items in the batch here
    }

    // After a batch is processed
    _argumentsBatch = [];
    _isBatchScheduled = false;
  }
}
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
You can use runtime logging to debug your custom function's XML manifest file or to look for errors in your custom functions in real time. Runtime logging is only available for Office 2016 desktop currently.

For full instructions on how to use runtime logging, [read this article](../testing/troubleshoot-manifest.md).

## Debugging
Development on debugging tools is still proceeding for custom functions, which are still in preview.  

At present, the best method for debugging is to use Excel through Office Online and use the F12 debugging tool native to your browser.  