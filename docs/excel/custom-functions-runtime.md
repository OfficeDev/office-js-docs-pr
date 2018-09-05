---
ms.date: 09/05/2018
description: 'Excel custom functions use a new JavaScript runtime, which differs from the standard Add-ins WebView control runtime.' 
title: 'Runtime for Excel Custom Function Add-ins'
---

# Runtime for Excel custom function add-ins

Excel custom functions do not use the standard Add-ins WebView control runtime, which behaves similarly to a browser. Instead, they employ a new JavaScript runtime.  

Advantages of this new runtime include:  
- **Speed**: Multiple instances of this runtime can work in parallel.
- **Network calls**: Through XHR, you can make individual or “streaming” batched requests to get data from the web.
- **Authentication**: AsyncStorage allows you to store tokens and set up authentication for users of your custom functions.
- **Modern Integration**: This runtime supports familiar technologies such as WebSockets and Dialog API.

> [!NOTE]
> This new JavaScript runtime does not provide access to the Document Object Model (DOM) or support libraries like JQuery which rely on the DOM.

The new runtime has two configurations: synchronous, for computation-only functions, and asynchronous, for functions which make network calls via XHR or utilize the other supported APIs, such as AsyncStorage.  

## Synchronous JavaScript runtime

This runtime can be parallelized, allowing for great improvements in speed and efficiency when performing calculations.

## Asynchronous JavaScript runtime

The asynchronous JavaScript runtime is optimized for asynchronous actions that custom functions may employ, like making network calls. The asynchronous JavaScript runtime supports four APIs, which you can use in your custom functions:  

- [XHR](#xhr)
- [WebSockets](#websockets)
- [AsyncStorage](#asyncstorage)
- [Dialog API](#dialog-api)

## XHR

XHR stands for XmlHttpRequest, which is a standard web API which performs HTTP requests, such as `POST`, `GET`, etc, to interact with servers. XHR in the new JavaScript runtime implements additional security measures by requiring Single Origin Policy and simple CORS.  

The sample below shows a function `getTemperature()`, which makes a call to the web to get the temperature of a particular area based on thermometer ID. XHR is used in the function `sendWebRequest()` make a `GET` request to an endpoint which can provide the data.  

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
        data.temperature=xhttp.responseText; //parsing would be needed here rather than blind assignment
      };
    xhttp.open("GET", "https://127.0.0.1:8080/temperature.aspx", true);
    xhttp.send();  
    }
}

```

## WebSockets

WebSockets is a networking protocol which creates real-time communication between a server and one or more clients. It is often used for chat applications because it allows you to read and write text simultaneously.  

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

## AsyncStorage

AsyncStorage is a key-value storage system that can be used to store authentication tokens or used for:

- Persistent
- Unencrypted
- Asynchronous
- Global to your Excel custom function

AsyncStorage can be used for storing authentication tokens and settings for reuse as well as a lightweight caching of information. Methods available on AsyncStorage include getItem, setItem, removeItem, clear, getAllKeys, flushGetRequests, multiGet, multiSet, and multiRemove. At this time, mergeItem and multiMerge are not supported methods.  

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

## Dialog API

The Dialog API allows you to require user authentication through an outside resource, such as Google or Facebook, before they can use your function. The Dialog API allows you to open a dialog box which prompts user sign-in.  

The code sample below illustrates the use of the Dialog API’s `displayWebDialog()` method.  

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

> [!NOTE]
> The Dialog API in the new JavaScript runtime differs from the [current Dialog API](../develop/dialog-api-in-office-add-ins.md) which works in the WebView control runtime.  