---
ms.date: 09/27/2018
description: Excel custom functions use a new JavaScript runtime, which differs from the standard Add-ins WebView control runtime.
title: Runtime for Excel custom functions
---

# Runtime for Excel custom functions (preview)

Custom functions employ a new JavaScript runtime, which differs the runtime utilized by other parts of an add-in, such as the task pane or other UI elements.

* This new runtime emphasizes fast performance for custom functions' calculations and has some [new APIs](#new-apis) that allow you to return data from external sources.

* On Excel for Windows, the new JavaScript runtime does not provide access to the Document Object Model (DOM) or support libraries like jQuery that rely on the DOM. If you would like your add-in to work seamlessly on both Excel Online and Excel for Windows, it is recommended that you avoid these libraries or accessing the DOM.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## New APIs

This new JavaScript runtime features two built-in APIs to facilitate some common web-based actions:

- Requesting external data with [fetch](#requesting-external-data) & [XHR](#xhr)
- [WebSockets](#websockets)

It also includes two APIs which can be used by both custom functions in this new runtime and by other parts of an add-in which do not use this new runtime, such as the task pane and UI elements.
- [AsyncStorage](#asyncstorage)
- [Dialog API](#dialog-api)

### Requesting external data

Requests for external data, such as information from the web, are supported in this new runtime. You can make XHR requests as detailed in the next section, or [use fetch](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API). In both cases, a new JavaScript Promise is returned.

>[!NOTE]
> While you may previously have had to specify `OfficeExtension.Promise` in add-ins to use Promises, as of September 2018, Promises are globally exposed.

#### XHR

XHR stands for [XmlHttpRequest](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers. In the new JavaScript runtime, XHR implements additional security measures by requiring [Same Origin Policy](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).  

In the following code sample, the `getTemperature()` function sends a web request to get the temperature of a particular area based on thermometer ID. The `sendWebRequest()` function uses XHR to issue a `GET` request to an endpoint that can provide the data.  

```js
function getTemperature(thermometerID) {
  return new Promise(function(setResult) {
      sendWebRequest(thermometerID, function(data){ //sendWebRequest is defined later in this code sample
          storeLastTemperature(thermometerID, data.temperature);
          setResult(data.temperature);
      });
  });
}

//Helper method that uses Office's implementation of XMLHttpRequest in the new JavaScript runtime for custom functions  
function sendWebRequest(thermometerID, data) {
    var xhttp = new XMLHttpRequest();
    xhttp.onreadystatechange = function() {
        if (this.readyState == 4 && this.status == 200) {
           data.temperature = JSON.parse(xhttp.responseText).temperature
          };
        xhttp.open("GET", "https://contoso.com/temperature/" + thermometerID), true)
        xhttp.send();  
    }
}

```

### WebSockets

[WebSockets](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) is a networking protocol that creates real-time communication between a server and one or more clients. It is often used for chat applications because it allows you to read and write text simultaneously.  

As shown in the following code sample, custom functions can use WebSockets. In this example, the WebSocket logs each message that it receives.

```typescript
const ws = new WebSocket('wss://bundles.office.com');
ws.onmessage = (message) => {
    console.log(`Received: ${message}`);
};
ws.onerror = (error) => {
    console.err(`Failed: ${error}`);
}
```

### AsyncStorage

AsyncStorage is a key-value storage system that can be used to store authentication tokens. It is:

- Persistent
- Unencrypted
- Asynchronous

AsyncStorage is globally available to all parts of your add-in. For custom functions, `AsyncStorage` is exposed as a global object. (For other parts of your add-in, such as task panes and other elements that use the WebView runtime, AsyncStorage is exposed through `OfficeRuntime`.) Each add-in has its own storage partition, with a default size of 5MB. 

The following methods are available on the `AsyncStorage` object:
 
 - `getItem`
 - `setItem`
 - `removeItem`
 - `clear`
 - `getAllKeys`
 - `flushGetRequests`
 - `multiGet`
 - `multiSet`
 - `multiRemove`

The following code sample calls the `AsyncStorage.getItem` function to retrieve a value from storage.

```typescript
_goGetData = async () => {
    try {
        const value = await AsyncStorage.getItem('toDoItem');
        if (value !== null) {
            //data exists and you can do something with it here
            }
        } catch (error) {
            //handle errors here
        }
    }
}
```

### Dialog API

The Dialog API enables you to open a dialog box that prompts user sign-in. You can use the Dialog API to require user authentication through an outside resource, such as Google or Facebook, before the user can use your function.

In the following code sample, the `getTokenViaDialog()` method uses the Dialog API’s `displayWebDialog()` method to open a dialog box.

```js
// Get auth token before calling my service, a hypothetical API that will deliver a stock price based on stock ticker string, such as "MSFT"

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
  
  //Helper
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
  }

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
        OfficeRuntime.displayWebDialog(url, {
          height: '50%',
          width: '50%',
          onMessage: function (message, dialog) {
            _cachedToken = message;
            resolve(message);
            dialog.closeDialog();
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
}
```

> [!NOTE]
> The Dialog API described in this section is part of the new JavaScript runtime. This API is different from the [Dialog API](../develop/dialog-api-in-office-add-ins.md) that can be used within task panes and add-in commands.

## See also

* [Create custom functions in Excel](custom-functions-overview.md)
* [Custom functions metadata](custom-functions-json.md)
* [Custom functions best practices](custom-functions-best-practices.md)
