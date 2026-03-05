---
ms.date: 03/03/2026
description: Request, stream, and cancel streaming of external data to your workbook with custom functions in Excel.
title: Receive and handle data with custom functions
ms.localizationpriority: medium
---

# Receive and handle data with custom functions

One of the ways that custom functions enhances Excel's power is by receiving data from locations other than the workbook, such as the web or a server (through [WebSockets](https://developer.mozilla.org/docs/Web/API/WebSockets_API)). You can request external data through an API like [`Fetch`](https://developer.mozilla.org/docs/Web/API/Fetch_API) or by using `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.

:::image type="content" source="../images/custom-functions-web-api.gif" alt-text="GIF of a custom function which streams the time from an API.":::

## Key points

- Return a JavaScript `Promise` from functions that fetch external data.
- Use streaming functions to continuously update cell values without user interaction.
- Streaming functions use the `@streaming` tag and `CustomFunctions.StreamingInvocation` parameter.
- The `onCanceled` callback handles cleanup when a function is canceled.
- WebSockets enable real-time data updates from servers with persistent connections.

## Functions that return data from external sources

Custom functions that retrieve data from external sources such as REST APIs or web services are asynchronous by nature. Excel needs to wait for the data to arrive before displaying results in the cell. To handle this, your function must:

1. Return a [JavaScript `Promise`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) to Excel.
2. Resolve the `Promise` with the final value using the callback function.

Excel automatically waits for the promise to resolve before displaying the result in the cell. This pattern works for one-time data requests. For continuous updates, use streaming functions instead.

### Fetch example

In the following code sample, the `webRequest` function reaches out to a hypothetical external API that tracks the number of people currently on the International Space Station. The function returns a JavaScript `Promise` and uses `fetch` to request information from the hypothetical API. The resulting data is transformed into JSON and the `names` property is converted into a string, which is used to resolve the promise.

When developing your own functions, consider performing an action if the web request doesn't complete in a timely manner or [batching up multiple API requests](custom-functions-batching.md).

```JS
/**
 * Requests the names of the people currently on the International Space Station.
 * Note: This function requests data from a hypothetical URL. In practice, replace the URL with a data source for your scenario.
 * @customfunction
 */
function webRequest() {
  let url = "https://www.contoso.com/NumberOfPeopleInSpace"; // This is a hypothetical URL.
  return new Promise(function (resolve, reject) {
    fetch(url)
      .then(function (response){
        return response.json();
        }
      )
      .then(function (json) {
        resolve(JSON.stringify(json.names));
      })
  })
}
```

> [!NOTE]
> Using `fetch` avoids nested callbacks and may be preferable to XHR in some cases.

### XHR example

In the following code sample, the `getStarCount` function calls the Github API to discover the amount of stars given to a particular user's repository. This is an asynchronous function which returns a JavaScript `Promise`. When data is obtained from the web call, the promise is resolved which returns the data to the cell.

```TS
/**
 * Gets the star count for a given Github organization or user and repository.
 * @customfunction
 * @param userName string name of organization or user.
 * @param repoName string name of the repository.
 * @return number of stars.
 */
async function getStarCount(userName: string, repoName: string) {

  const url = "https://api.github.com/repos/" + userName + "/" + repoName;

  let xhttp = new XMLHttpRequest();

  return new Promise(function(resolve, reject) {
    xhttp.onreadystatechange = function() {
      if (xhttp.readyState !== 4) return;

      if (xhttp.status == 200) {
        resolve(JSON.parse(xhttp.responseText).watchers_count);
      } else {
        reject({
          status: xhttp.status,

          statusText: xhttp.statusText
        });
      }
    };

    xhttp.open("GET", url, true);

    xhttp.send();
  });
}
```

## Make a streaming function

Streaming custom functions enable you to output data to cells that updates repeatedly, without requiring a user to explicitly refresh anything. This is useful for displaying live data from services, such as stock prices, sensor readings, or real-time analytics, like the function in [the custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md).

To declare a streaming function, you can use either of the following two options.

- The `@streaming` JSDoc tag.
- The `CustomFunctions.StreamingInvocation` invocation parameter.

Streaming functions differ from regular asynchronous functions in that they can call `setResult` multiple times to update the cell value continuously, rather than returning a single result.

### Basic streaming example

The following code sample is a custom function that adds a number to the result every second. Note the following about this code.

- Excel displays each new value automatically using the `setResult` method.
- The second input parameter, `invocation`, isn't displayed to end users in Excel when they select the function from the autocomplete menu.
- The `onCanceled` callback defines the function that runs when the function is canceled.
- Streaming isn't necessarily tied to making a web request. In this case, the function isn't making a web request but is still getting data at set intervals, so it requires the use of the streaming `invocation` parameter.

```JS
/**
 * Increments a value once a second.
 * @customfunction INC increment
 * @param {number} incrementBy Amount to increment.
 * @param {CustomFunctions.StreamingInvocation<number>} invocation
 */
function increment(incrementBy, invocation) {
  let result = 0;
  const timer = setInterval(() => {
    result += incrementBy;
    invocation.setResult(result);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}
```

### Streaming data from a web service

The following example shows a streaming function that fetches stock prices from a web service every 10 seconds.

```JS
/**
 * Streams stock price updates.
 * @customfunction
 * @param {string} ticker Stock ticker symbol.
 * @param {CustomFunctions.StreamingInvocation<number>} invocation
 */
function stockPrice(ticker, invocation) {
  const updateInterval = 10000; // Update every 10 seconds.

  const timer = setInterval(() => {
    // Replace with your actual API endpoint.
    fetch(`https://api.example.com/stock/${ticker}`)
      .then(response => response.json())
      .then(data => {
        invocation.setResult(data.price);
      })
      .catch(error => {
        // Return the #N/A error if stock price is unavailable.
        invocation.setResult(
          new CustomFunctions.Error(CustomFunctions.ErrorCode.notAvailable)
        );
      });
  }, updateInterval);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}
```

> [!NOTE]
> For an example of how to return a dynamic spill array from a streaming function, see [Return multiple results from your custom function: Code samples](custom-functions-dynamic-arrays.md#code-samples).

## Cancel a function

Excel automatically cancels the execution of a function in the following situations.

- When the user edits or deletes a cell that references the function.
- When one of the arguments (inputs) for the function changes. In this case, a new function call is triggered following the cancellation.
- When the user triggers recalculation manually. In this case, a new function call is triggered following the cancellation.

Proper cleanup in the `onCanceled` callback is important to prevent unnecessary network requests. Always clear timers, close connections, and abort pending requests when a function is canceled. You can also consider setting a default streaming value to handle cases when a request is made but you are offline.

> [!NOTE]
> There is also a category of functions called cancelable functions which use the `@cancelable` JSDoc tag. Cancelable functions allow a web request to be terminated in the middle of the request.
>
> A streaming function can't use the `@cancelable` tag, but streaming functions can include an `onCanceled` callback function. Only asynchronous custom functions which return one value can use the `@cancelable` JSDoc tag. See [Autogenerate JSON metadata: @cancelable](custom-functions-json-autogeneration.md#cancelable) to learn more about the `@cancelable` tag.

### Use an invocation parameter

The `invocation` parameter is the last parameter of any custom function by default. The `invocation` parameter gives context about the cell (such as its address and contents) and allows you to use the `setResult` method and `onCanceled` event to define what a function does when it streams (`setResult`) or is canceled (`onCanceled`).

The invocation handler needs to be of type [`CustomFunctions.StreamingInvocation`](/javascript/api/custom-functions-runtime/customfunctions.streaminginvocation) or [`CustomFunctions.CancelableInvocation`](/javascript/api/custom-functions-runtime/customfunctions.cancelableinvocation) to process web requests.

See [Invocation parameter](custom-functions-parameter-options.md#invocation-parameter) to learn about other potential uses of the `invocation` argument and how it corresponds with the [Invocation](/javascript/api/custom-functions-runtime/customfunctions.invocation) object.

## Receiving data via WebSockets

Within a custom function, you can use [WebSockets](https://developer.mozilla.org/docs/Web/API/WebSockets_API) to exchange data over a persistent connection with a server. WebSockets are useful for real-time data that updates frequently, such as financial tickers, chat messages, or sensor data. Using WebSockets, your custom function can open a connection with a server and then automatically receive messages from the server when certain events occur, without having to explicitly poll the server for data.

### WebSocket streaming example

The following code sample shows a streaming function that uses WebSockets to receive real-time updates.

```js
/**
 * Streams real-time data via WebSocket.
 * @customfunction
 * @param {string} symbol Data symbol to monitor.
 * @param {CustomFunctions.StreamingInvocation<string>} invocation
 */
function streamWebSocket(symbol, invocation) {
  const ws = new WebSocket('wss://example.com/data');

  ws.onopen = () => {
    // Subscribe to updates for the specified symbol.
    ws.send(JSON.stringify({ subscribe: symbol }));
  };

  ws.onmessage = (event) => {
    const data = JSON.parse(event.data);
    invocation.setResult(data.value);
  };

  ws.onerror = (error) => {
    // Return the #N/A error if connection fails.
    invocation.setResult(
      new CustomFunctions.Error(CustomFunctions.ErrorCode.notAvailable)
    );
  };

  invocation.onCanceled = () => {
    ws.close();
  };
}
```

## Next steps

- Learn about [different parameter types your functions can use](custom-functions-parameter-options.md).
- Discover how to [batch multiple API calls](custom-functions-batching.md).

## See also

- [Excel custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md)
- [Custom functions parameter options](custom-functions-parameter-options.md)
- [Batch custom function calls for a remote service](custom-functions-batching.md)
- [Return multiple results from your custom function](custom-functions-dynamic-arrays.md)
- [Volatile values in functions](custom-functions-volatile.md)
- [Create JSON metadata for custom functions](custom-functions-json-autogeneration.md)
- [Create custom functions in Excel](custom-functions-overview.md)
