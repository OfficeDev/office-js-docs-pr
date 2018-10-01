---
ms.date: 09/27/2018
description: Excel custom functions use a new JavaScript runtime, which differs from the standard Add-ins WebView control runtime.
title: Runtime for Excel custom functions
---

# Runtime for Excel custom functions (preview)

Custom functions use a new JavaScript runtime, which differs the runtime used by other parts of an add-in, such as the task pane or other UI elements. This JavaScript runtime is designed to optimize performance of calculations in custom functions and exposes some new APIs. Some of these APIs allow you to perform common web-based actions, such as requesting external data or enabling chat. Others, like `AsyncStorage` have been developed so they are common to both custom functions and other parts of an add-in.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## Requesting external data using fetch and XHR

Requests for external data, such as information from the web, are supported in this JavaScript runtime through two APIs:

* Fetch: you can [use fetch](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) to request data

* XHR: you can make an XHR request, as detailed below. 

XHR stands for [XmlHttpRequest](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers. In this JavaScript runtime, XHR implements additional security measures by requiring [Same Origin Policy](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).  

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

//Helper method that uses Office's implementation of XMLHttpRequest in the JavaScript runtime for custom functions  
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

When using fetch or XHR, a new JavaScript Promise is returned. You'll note that you may have had to specify `OfficeExtension.Promise` in the past to use Promises, but as of September 2018, they are globally exposed.

## Enabling chat using WebSockets

[WebSockets](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) are also globally exposed in the JavaScript runtime. Websockets is a networking protocol that creates real-time communication between a server and one or more clients. It is often used for chat applications because it allows you to read and write text simultaneously. 

This sample below shows how you can declare a websocket which logs each message it receives. 

```typescript
const ws = new WebSocket('wss://bundles.office.com');
ws.onmessage = (message) => {
    console.log(`Received: ${message}`);
};
ws.onerror = (error) => {
    console.err(`Failed: ${error}`);
}
```

## Managing data with AsyncStorage

AsyncStorage is a persistent, unencrypted, key-value storage system that can be used to store data. It can be used as an alternative to localStorage, which is not available to custom functions. AsyncStorage is accessible to custom functions as a global object and to all other parts of your add-in through `OfficeRuntime.AsyncStorage`. Each add-in has a 5MB storage partition by default.

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
```

## Enabling pop-up dialog windows through Dialog API

The Dialog API enables you to open a pop-up window which a user can interact with. A common usage scenario for the Dialog API is to prompt user sign-in for authentication via an outside resource, but it can be used for any situation which calls for a dialog box. It can be used by both custom functions and all other parts of your add-in, as it is exposed through `OfficeRuntime.displayWebDialogOptions`. Note that this API is different from the [Dialog API](../develop/dialog-api-in-office-add-ins.md) that can only be used within task panes and add-in commands.

In the following code sample, the function `getTokenViaDialog` uses the Dialog API’s `displayWebDialogOptions` method to open a dialog box.

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
        OfficeRuntime.displayWebDialogOptions(url, {
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

## Best practices

To develop an add-in which works both on Excel for Windows and Excel Online, you should avoid accessing the Document Object Model (DOM) or support libraries like jQuery which rely on the DOM. This is because Excel Online uses and renders everything in the browser, whereas Excel for Windows uses the JavaScript runtime, which does not support the DOM and libraries which use it.

## See also

* [Create custom functions in Excel](custom-functions-overview.md)
* [Custom functions metadata](custom-functions-json.md)
* [Custom functions best practices](custom-functions-best-practices.md)
