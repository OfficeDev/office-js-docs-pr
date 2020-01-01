---
ms.date: 12/31/2019
description: Request, stream, and cancel streaming of external data to your workbook with custom functions in Excel
title: Receive and handle data with custom functions
localization_priority: Priority
---

# Receive and handle data with custom functions

One of the ways that custom functions enhances Excel's power is by receiving data from locations other than the workbook, such as the web or a server (through WebSockets). You can request external data through an API like [`Fetch`](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) or by using `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

![A gif of a custom function which streams the time from an API](../images/WebApi2.gif)

## Functions that return data from external sources

If a custom function retrieves data from an external source such as the web, it must:

1. Return a JavaScript Promise to Excel.
2. Resolve the Promise with the final value using the callback function.

### Fetch example

In the following code sample, the **webRequest** function reaches out to the hypothetical Contoso "Number of People in Space" API, which tracks the number of people currently on the International Space Station. The function returns a JavaScript Promise and uses fetch to request information from the API. The resulting data is transformed into JSON and the `names` property is converted into a string, which is used to resolve the Promise.

When developing your own functions, you may want to perform an action if the web request does not complete in a timely manner or consider [batching up multiple API requests](./custom-functions-batching.md).

```JS
/**
 * Requests the names of the people currently on the International Space Station from a hypothetical API.
 * @customfunction
 */
function webRequest() {
  let url = "https://www.contoso.com/NumberOfPeopleInSpace";
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

>[!NOTE]
>Using `Fetch` avoids nested callbacks and may be preferable to XHR in some cases.

### XHR example

Within custom functions runtime, XHR implements additional security measures by requiring [Same Origin Policy](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).

Note that a simple CORS implementation cannot use cookies and only supports simple methods (GET, HEAD, POST). Simple CORS accepts simple headers with field names `Accept`, `Accept-Language`, `Content-Language`. You can also use a Content-Type header in simple CORS, provided that the content type is `application/x-www-form-urlencoded`, `text/plain`, or `multipart/form-data`.

In the following code sample, the **getStarCount** function calls the Github API to discover the amount of stars given to a particular user's repository. This is an asynchronous function which returns a JavaScript Promise. When data is obtained from the web call, the Promise is resolved which returns the data to the cell.

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

Streaming custom functions enable you to output data to cells that updates repeatedly, without requiring a user to explicitly refresh anything. This can be useful to check live data from a service online, like the function in [the custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md).

To declare a streaming function, either use the tag `@streaming` or make use of the `CustomFunctions.StreamingInvocation` invocation parameter, which will indicate that your function is streaming. To alert users to the fact that your function may re-evaluate based on new information, consider putting stream or other wording to indicate this in the name or description of your function.

The following code sample is a custom function that adds a number to the result every second. Note the following about this code:

- Excel displays each new value automatically using the `setResult` method.
- The second input parameter, invocation, is not displayed to end users in Excel when they select the function from the autocomplete menu.
- The `onCanceled` callback defines the function that executes when the function is canceled.
- Streaming isn't necessarily tied to making a web request: in this case, the function isn't making a web request but is still getting data at set intervals, so it requires the use of the streaming `invocation` parameter.

```js
/**
 * Increments a value once a second.
 * @customfunction INC increment
 * @param {number} incrementBy Amount to increment
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

In addition to knowing about the `onCanceled` callback, you should also know that Excel cancels the execution of a function in the following situations:

- When the user edits or deletes a cell that references the function.
- When one of the arguments (inputs) for the function changes. In this case, a new function call is triggered following the cancellation.
- When the user triggers recalculation manually. In this case, a new function call is triggered following the cancellation.

You can also consider setting a default streaming value to handle cases when a request is made but you are offline.

> [!NOTE]
> Note that there are also a category of functions called cancelable functions, which are _not_ related to streaming functions. Previous versions of custom functions required you to declare `"cancelable": true` and `"streaming": true` in JSON written by hand. Since the introduction of autogenerated metadata, only asynchronous custom functions which return one value are cancelable. Cancelable functions allow a web request to be terminated in the middle of a request, using a [`CancelableInvocation`](/javascript/api/custom-functions-runtime/customfunctions.cancelableinvocation) to decide what to do upon cancellation. Declare a cancelable function using the tag `@cancelable`.

### Using an invocation parameter

The `invocation` parameter is the last parameter of any custom function by default. The `invocation` parameter gives context about the cell (such as its address and contents) and also allows you to use `setResult` and `onCanceled` methods. These methods define what a function does when the function streams (`setResult`) or is canceled (`onCanceled`).

If you're using TypeScript, the invocation handler needs to be of type `CustomFunctions.StreamingInvocation` or `CustomFunctions.CancelableInvocation`.

## Receive data via WebSockets

Within a custom function, you can use WebSockets to exchange data over a persistent connection with a server. By using WebSockets, your custom function can open a connection with a server and then automatically receive messages from the server when certain events occur, without having to explicitly poll the server for data.

### WebSockets example

The following code sample establishes a WebSocket connection and then logs each incoming message from the server.

```js
let ws = new WebSocket('wss://bundles.office.com');

ws.onmessage(message) {
    console.log(`Received: ${message}`);
}

ws.onerror(error){
    console.err(`Failed: ${error}`);
}
```

## Next steps

- Learn about [different parameter types your functions can use](custom-functions-parameter-options.md).
- Discover how to [batch multiple API calls](custom-functions-batching.md).

## See also

- [Volatile values in functions](custom-functions-volatile.md)
- [Create JSON metadata for custom functions](custom-functions-json-autogeneration.md)
- [Custom functions metadata](custom-functions-json.md)
- [Runtime for Excel custom functions](custom-functions-runtime.md)
- [Create custom functions in Excel](custom-functions-overview.md)
- [Excel custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md)
