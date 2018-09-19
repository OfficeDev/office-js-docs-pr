---
ms.date: 09/18/2018
description: Excel custom functions use a new JavaScript runtime, which differs from the standard Add-ins WebView control runtime.
title: Runtime for Excel Custom Function Add-ins
---

# Runtime for Excel custom function add-ins

Custom functions extend Excel’s capabilities using a new JavaScript runtime. This runtime utilizes a sandboxed JavaScript engine rather than a web browser. Additionally, it prioritizes performance, allowing you to run thousands of custom functions simultaneously.  

> [!NOTE]
> The code for your add-in may include other parts, such as task panes and other UI elements.  
> These will continue to run in the browser-like WebView runtime that you are used to. The new runtime only applies to the custom functions related code in your add-in.  

## Differences between WebView runtime and the new JavaScript runtime

- The new JavaScript runtime used by custom functions does not provide access to the Document Object Model (DOM) or support libraries like JQuery which rely on the DOM.
- In the JavaScript file which defines your functions (if using yo office, **customfunctions.js**), you can now return a regular JavaScript `Promise` instead of using `OfficeExtension.Promise`.
- In the JSON file which describes your functions (if using yo office, **customfunctions.json**), you no longer need to specify “sync” or “async” under “options”.  

## New APIs 
The new JavaScript runtime utilized by custom functions has four APIs:

- [XHR](#xhr)
- [WebSockets](#websockets)
- [AsyncStorage](#asyncstorage)
- [Dialog API](#dialog-api)

### XHR

XHR stands for [XmlHttpRequest](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), which is a standard web API which performs HTTP requests, such as `POST`, `GET`, etc, to interact with servers. XHR in the new JavaScript runtime implements additional security measures by requiring [Same Origin Policy](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).  

The sample below shows a function `getTemperature()`, which makes a call to the web to get the temperature of a particular area based on thermometer ID. XHR is used in the function `sendWebRequest()` to make a `GET` request to an endpoint which can provide the data.  

```js
function getTemperature(thermometerID) {
  return new Promise(function(setResult) {
      sendWebRequest(thermometerID, function(data){ //sendWebRequest utilizes XHR, see its definition below
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
        data.temperature = xhttp.responseText; //parsing is needed here rather than blind assignment
      };
    xhttp.open("GET", "https://127.0.0.1:8080/temperature.aspx", true);
    xhttp.send();  
    }
}

```

### WebSockets

[WebSockets](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) is a networking protocol which creates real-time communication between a server and one or more clients. It is often used for chat applications because it allows you to read and write text simultaneously.  

As you can see in the sample below, custom functions can use WebSockets. In this example, the WebSocket logs any message sent to it.  

```js
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

Additionally, AsyncStorage is available globally to all parts of your add-in. For custom functions, AsyncStorage is exposed as a global object. For other parts of your add-in, such as task panes and other elements which utilize the typical WebView runtime, AsyncStorage is exposed through `Office.Runtime`.

 Methods available on AsyncStorage include getItem, setItem, removeItem, clear, getAllKeys, flushGetRequests, multiGet, multiSet, and multiRemove. At this time, mergeItem and multiMerge are not supported methods.

Each add-in has its own storage partition, with a default of 5MB of storage.  

The code sample below illustrates the process of getting an item from AsyncStorage:

```js
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

The Dialog API allows you to require user authentication through an outside resource, such as Google or Facebook, before they can use your function. The Dialog API enables you to open a dialog box which prompts user sign-in.  

The code sample below illustrates the use of the Dialog API’s `displayWebDialog()` method.  

```js
// Get auth token before calling my service, a hypothetical API which will deliver a stock price based on stock ticker string, such as "MSFT"
async function getStock(ticker) {
    const token = await getToken();
    const data = await (await fetch(https://myservice.com/?token=token&ticker= + ticker).json());
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
           height: 50,
           width: 50%,
           hideTitle: true,
           onMessage: (message, dialog) => {
               const json = JSON.parse(message);
                    if (json.type === "token_succeeded") {
                        resolve(json.value);
                        dialog.closeDialog();
                        return;
                    }
            // Otherwise, handle other messages.
           },
           onClose: () => reject("User closed dialog"),
        }).catch(e => reject(e));
    });
}
```

> [!NOTE]
> The Dialog API described in this section is part of the new JavaScript runtime for custom functions and can be used only within custom functions. This API is different from the [Dialog API](../develop/dialog-api-in-office-add-ins.md) that can be used within task panes and add-in commands.