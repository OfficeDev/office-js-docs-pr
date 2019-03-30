---
ms.date: 03/21/2019
description: Request, stream, and cancel streaming of external data to your workbook with custom functions in Excel
title: Web requests and other data handling with custom functions (preview)
localization_priority: Priority
---

# Receiving and handling data with custom functions

One of the ways that custom functions enhance Excel's power is by receiving data from locations other than the workbook, such as the web or a server (through WebSockets). Custom functions can request data through XHR and fetch requests as well as stream this data in real time.

The documentation below illustrates some samples of web requests, but to build a streaming function for yourself, try the [Custom functions tutorial](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows).

## Functions that return data from external sources

If a custom function retrieves data from an external source such as the web, it must:

1. Return a JavaScript Promise to Excel.
2. Resolve the Promise with the final value using the callback function.

You can request external data through an API like [`Fetch`](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) or by using `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.

Within custom functions runtime, XHR implements additional security measures by requiring [Same Origin Policy](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).

Note that a simple CORS implementation cannot use cookies and only supports simple methods (GET, HEAD, POST). Simple CORS accepts simple headers with field names `Accept`, `Accept-Language`, `Content-Language`. You can also use a Content-Type header in simple CORS, provided that the content type is `application/x-www-form-urlencoded`, `text/plain`, or `multipart/form-data`.

### XHR example

In the following code sample, the **getTemperature** function calls the sendWebRequest function to get the temperature of a particular area based on thermometer ID. The sendWebRequest function uses XHR to issue a GET request to an endpoint that can provide the data.

```JavaScript
function getTemperature(thermometerID) {
  return new Promise(function(setResult) {
      sendWebRequest(thermometerID, function(data){ 
          storeLastTemperature(thermometerID, data.temperature);
          setResult(data.temperature);
      });
  });
}

// Helper method that uses Office's implementation of XMLHttpRequest in the JavaScript runtime for custom functions  
function sendWebRequest(thermometerID, data) {
    var xhttp = new XMLHttpRequest();
    xhttp.onreadystatechange = function() {
        if (this.readyState == 4 && this.status == 200) {
           data.temperature = JSON.parse(xhttp.responseText).temperature
        };

        //set Content-Type to application/text. Application/json is not currently supported with Simple CORS
        xhttp.setRequestHeader("Content-Type", "application/text");
        xhttp.open("GET", "https://contoso.com/temperature/" + thermometerID), true)
        xhttp.send();  
    }
}

CustomFunctions.associate("GETTEMPERATURE", getTemperature);
```

For another sample of an XHR request with more context, see the `getFile` function within [this file](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload/blob/master/FileDownloadSampleWeb/Home.js) in the [Office-Add-in-JavaScript-FileDownload](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload) Github repository.

### Fetch example

In the following code sample, the stockPriceStream function uses a stock ticker symbol to get the price of a stock every 1000 milliseconds. For more details about this sample and to get the accompanying JSON, see the [Custom functions tutorial](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows#create-a-streaming-asynchronous-custom-function). 

```JavaScript
function stockPriceStream(ticker, handler) {
    var updateFrequency = 1000 /* milliseconds*/;
    var isPending = false;

    var timer = setInterval(function() {
        // If there is already a pending request, skip this iteration:
        if (isPending) {
            return;
        }

        var url = "https://api.iextrading.com/1.0/stock/" + ticker + "/price";
        isPending = true;

        fetch(url)
            .then(function(response) {
                return response.text();
            })
            .then(function(text) {
                handler.setResult(parseFloat(text));
            })
            .catch(function(error) {
                handler.setResult(error);
            })
            .then(function() {
                isPending = false;
            });
    }, updateFrequency);

    handler.onCanceled = () => {
        clearInterval(timer);
    };
}

CustomFunctions.associate("STOCKPRICESTREAM", stockPriceStream);
```

## Receiving data via WebSockets

Within a custom function, you can use WebSockets to exchange data over a persistent connection with a server. By using WebSockets, your custom function can open a connection with a server and then automatically receive messages from the server when certain events occur, without having to explicitly poll the server for data.

### WebSockets example

The following code sample establishes a WebSocket connection and then logs each incoming message from the server.

```JavaScript
var ws = new WebSocket('wss://bundles.office.com');

ws.onmessage(message) {
    console.log(`Recieved: ${message}`);
}

ws.onerror(error){
    console.err(`Failed: ${error}`);
}
```

## Streaming and cancelable functions

Streaming custom functions enable you to output data to cells repeatedly over time, without requiring a user to explicitly request data refresh.

Cancelable custom functions enable you to cancel the execution of a streaming custom function to reduce its bandwidth consumption, working memory, and CPU load. Excel automatically cancels the execution of a function in the following situations:

- When the user edits or deletes a cell that references the function.
- When one of the arguments (inputs) for the function changes. In this case, a new function call is triggered following the cancellation.
- When the user triggers recalculation manually. In this case, a new function call is triggered following the cancellation.

To declare a streaming or cancelable function, you will need to do two things: 

- Add streaming and/or cancelable properties to the function's JSON metadata, by declaring `"streaming": true` or `"cancelable": true`. 
- Declare an invocation parameter which carries important context of the cell being called. 

### Using an invocation parameter

When setting an invocation parameter, it should come last in the list of parameters specified. An invocation will allow you to use `setResult` on streaming functions or `onCanceled` methods on cancelable functions. 

If using Typescript, note that the invocation handler you use will need to be specified by type `CustomFunctions.StreamingInvocation` or `CustomFunctions.CancelableInvocation`. 

### Streaming and cancelable function example
The following code sample is a custom function that adds a number to the result every second. Note the following about this code:

- Excel displays each new value automatically using the `setResult` method.
- The second input parameter, invocation, is not displayed to end users in Excel when they select the function from the autocomplete menu. 
- The `onCanceled` callback defines the function that executes when the function is canceled. 

```JavaScript
function incrementValue(increment, invocation){
  var result = 0;
  setInterval(function(){
    result += increment;
    invocation.setResult(result);
  }, 1000);

  invocation.onCanceled = function(){
    clearInterval(timer);
  }
}

CustomFunctions.associate("INCREMENTVALUE", incrementValue);
```
The corresponding JSON for this function is as follows. Note that the invocation parameter does not need to be declared in the JSON file. 

```JSON
{
  "id": "INCREMENTVALUE",
  "name": "INCREMENTVALUE",
  "description": "Periodically increment a value",
  "helpUrl": "http://www.contoso.com",
  "result": {
    "type": "number",
    "dimensionality": "scalar"
  },
  "parameters": [
    {
      "name": "increment",
      "description": "Amount to increment",
      "type": "number",
      "dimensionality": "scalar"
    }
  ],
  "options": {
    "cancelable": true,
    "stream": true
  }
}
```

## See also

* [Excel custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md)
* [Custom functions metadata](custom-functions-json.md)
* [Runtime for Excel custom functions](custom-functions-runtime.md)
* [Custom functions best practices](custom-functions-best-practices.md)
* [Custom functions changelog](custom-functions-changelog.md)
