---
ms.date: 09/09/2022
description: Batch custom functions together to reduce network calls to a remote service.
title: Batching custom function calls for a remote service
ms.topic: best-practice
ms.localizationpriority: medium
---

# Batch custom function calls for a remote service

If your custom functions call a remote service you can use a batching pattern to reduce the number of network calls to the remote service. To reduce network round trips you batch all the calls into a single call to the web service. This is ideal when the spreadsheet is recalculated.

For example, if someone used your custom function in 100 cells in a spreadsheet, and then recalculated the spreadsheet, your custom function would run 100 times and make 100 network calls. By using a batching pattern, the calls can be combined to make all 100 calculations in a single network call.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## View the completed sample

To view the completed sample, follow this article and paste the code examples into your own project. For example, to create a new custom function project for TypeScript use the [Yeoman generator for Office Add-ins](../develop/yeoman-generator-overview.md), then add all the code from this article to the project. Run the code and try it out.

Alternatively, download or view the complete sample project at [Custom function batching pattern](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Excel-custom-functions/Batching). If you want to view the code in whole before reading any further, take a look at the [script file](https://github.com/OfficeDev/Office-Add-in-samples/blob/main/Excel-custom-functions/Batching/src/functions/functions.js).

## Create the batching pattern in this article

To set up batching for your custom functions you'll need to write three main sections of code.

1. A [push operation](#add-the-_pushoperation-function) to add a new operation to the batch of calls each time Excel calls your custom function.
2. A [function to make the remote request](#make-the-remote-request) when the batch is ready.
3. [Server code to respond to the batch request](#process-the-batch-call-on-the-remote-service), calculate all of the operation results, and return the values.

In the following sections, you'll learn how to construct the code one example at a time. It's recommended you create a brand-new custom functions project using the [Yeoman generator for Office Add-ins](../develop/yeoman-generator-overview.md) generator. To create a new project, see [Get started developing Excel custom functions](../quickstarts/excel-custom-functions-quickstart.md). You can use TypeScript or JavaScript.

## Batch each call to your custom function

Your custom functions work by calling a remote service to perform the operation and calculate the result they need. This provides a way for them to store each requested operation into a batch. Later you'll see how to create a `_pushOperation` function to batch the operations. First, take a look at the following code example to see how to call `_pushOperation` from your custom function.

In the following code, the custom function performs division but relies on a remote service to do the actual calculation. It calls `_pushOperation` to batch the operation along with other operations to the remote service. It names the operation **div2**. You can use any naming scheme you want for operations as long as the remote service is also using the same scheme (more on the remote service later). Also, the arguments the remote service will need to run the operation are passed.

### Add the div2 custom function

Add the following code to your **functions.js** or **functions.ts** file (depending on if you used JavaScript or TypeScript).

```javascript
/**
 * Divides two numbers using batching
 * @CustomFunction
 * @param dividend The number being divided
 * @param divisor The number the dividend is divided by
 * @returns The result of dividing the two numbers
 */
function div2(dividend, divisor) {
  return _pushOperation("div2", [dividend, divisor]);
}
```

### Add global variables for tracking batch requests

Next, add two global variables to your **functions.js** or **functions.ts** file. `_isBatchedRequestScheduled` is important later for timing batch calls to the remote service.

```javascript
let _batch = [];
let _isBatchedRequestScheduled = false;
```

### Add the `_pushOperation` function

When Excel calls your custom function, you need to push the operation into the batch array. The following **_pushOperation** function code shows how to add a new operation from a custom function. It creates a new batch entry, creates a new promise to resolve or reject the operation, and pushes the entry into the batch array.

This code also checks to see if a batch is scheduled. In this example, each batch is scheduled to run every 100ms. You can adjust this value as needed. Higher values result in bigger batches being sent to the remote service, and a longer wait time for the user to see results. Lower values tend to send more batches to the remote service, but with a quick response time for users.

The function creates an **invocationEntry** object that contains the string name of which operation to run. For example, if you had two custom functions named `multiply` and `divide`, you could reuse those as the operation names in your batch entries. `args` holds the arguments that were passed to your custom function from Excel. And finally, `resolve` or `reject` methods store a promise holding the information the remote service returns.

Add the following code to your **functions.js** or **functions.ts** file.

```javascript
// This function encloses your custom functions as individual entries,
// which have some additional properties so you can keep track of whether or not
// a request has been resolved or rejected.
function _pushOperation(op, args) {
  // Create an entry for your custom function.
  console.log("pushOperation");
  const invocationEntry = {
    operation: op, // e.g., sum
    args: args,
    resolve: undefined,
    reject: undefined,
  };

  // Create a unique promise for this invocation,
  // and save its resolve and reject functions into the invocation entry.
  const promise = new Promise((resolve, reject) => {
    invocationEntry.resolve = resolve;
    invocationEntry.reject = reject;
  });

  // Push the invocation entry into the next batch.
  _batch.push(invocationEntry);

  // If a remote request hasn't been scheduled yet,
  // schedule it after a certain timeout, e.g., 100 ms.
  if (!_isBatchedRequestScheduled) {
    console.log("schedule remote request");
    _isBatchedRequestScheduled = true;
    setTimeout(_makeRemoteRequest, 100);
  }

  // Return the promise for this invocation.
  return promise;
}
```

## Make the remote request

The purpose of the `_makeRemoteRequest` function is to pass the batch of operations to the remote service, and then return the results to each custom function. It first creates a copy of the batch array. This allows concurrent custom function calls from Excel to immediately begin batching in a new array. The copy is then turned into a simpler array that does not contain the promise information. It wouldn't make sense to pass the promises to a remote service since they would not work. The `_makeRemoteRequest` will either reject or resolve each promise based on what the remote service returns.

Add the following code to your **functions.js** or **functions.ts** file.

```javascript
// This is a private helper function, used only within your custom function add-in.
// You wouldn't call _makeRemoteRequest in Excel, for example.
// This function makes a request for remote processing of the whole batch,
// and matches the response batch to the request batch.
function _makeRemoteRequest() {
  // Copy the shared batch and allow the building of a new batch while you are waiting for a response.
  // Note the use of "splice" rather than "slice", which will modify the original _batch array
  // to empty it out.
  try{
  console.log("makeRemoteRequest");
  const batchCopy = _batch.splice(0, _batch.length);
  _isBatchedRequestScheduled = false;

  // Build a simpler request batch that only contains the arguments for each invocation.
  const requestBatch = batchCopy.map((item) => {
    return { operation: item.operation, args: item.args };
  });
  console.log("makeRemoteRequest2");
  // Make the remote request.
  _fetchFromRemoteService(requestBatch)
    .then((responseBatch) => {
      console.log("responseBatch in fetchFromRemoteService");
      // Match each value from the response batch to its corresponding invocation entry from the request batch,
      // and resolve the invocation promise with its corresponding response value.
      responseBatch.forEach((response, index) => {
        if (response.error) {
          batchCopy[index].reject(new Error(response.error));
          console.log("rejecting promise");
        } else {
          console.log("fulfilling promise");
          console.log(response);

          batchCopy[index].resolve(response.result);
        }
      });
    });
    console.log("makeRemoteRequest3");
  } catch (error) {
    console.log("error name:" + error.name);
    console.log("error message:" + error.message);
    console.log(error);
  }
}
```

### Modify `_makeRemoteRequest` for your own solution

The `_makeRemoteRequest` function calls `_fetchFromRemoteService` which, as you'll see later, is just a mock representing the remote service. This makes it easier to study and run the code in this article. But when you want to use this code for an actual remote service you should make the following changes.

- Decide how to serialize the batch operations over the network. For example, you may want to put the array into a JSON body.
- Instead of calling `_fetchFromRemoteService` you need to make the actual network call to the remote service passing the batch of operations.

## Process the batch call on the remote service

The last step is to handle the batch call in the remote service. The following code sample shows the `_fetchFromRemoteService` function. This function unpacks each operation, performs the specified operation, and returns the results. For learning purposes in this article, the `_fetchFromRemoteService` function is designed to run in your web add-in and mock a remote service. You can add this code to your **functions.js** or **functions.ts** file so that you can study and run all the code in this article without having to set up an actual remote service.

Add the following code to your **functions.js** or **functions.ts** file.

```javascript
// This function simulates the work of a remote service. Because each service
// differs, you will need to modify this function appropriately to work with the service you are using. 
// This function takes a batch of argument sets and returns a promise that may contain a batch of values.
// NOTE: When implementing this function on a server, also apply an appropriate authentication mechanism
//       to ensure only the correct callers can access it.
async function _fetchFromRemoteService(requestBatch) {
  // Simulate a slow network request to the server.
  console.log("_fetchFromRemoteService");
  await pause(1000);
  console.log("postpause");
  return requestBatch.map((request) => {
    console.log("requestBatch server side");
    const { operation, args } = request;

    try {
      if (operation === "div2") {
        // Divide the first argument by the second argument.
        return {
          result: args[0] / args[1]
        };
      } else if (operation === "mul2") {
        // Multiply the arguments for the given entry.
        const myResult = args[0] * args[1];
        console.log(myResult);
        return {
          result: myResult
        };
      } else {
        return {
          error: `Operation not supported: ${operation}`
        };
      }
    } catch (error) {
      return {
        error: `Operation failed: ${operation}`
      };
    }
  });
}

function pause(ms) {
  console.log("pause");
  return new Promise((resolve) => setTimeout(resolve, ms));
}
```

### Modify `_fetchFromRemoteService` for your live remote service

To modify the `_fetchFromRemoteService` function to run in your live remote service, make the following changes.

- Depending on your server platform (Node.js or others) map the client network call to this function.
- Remove the `pause` function which simulates network latency as part of the mock.
- Modify the function declaration to work with the parameter passed if the parameter is changed for network purposes. For example, instead of an array, it may be a JSON body of batched operations to process.
- Modify the function to perform the operations (or call functions that do the operations).
- Apply an appropriate authentication mechanism. Ensure that only the correct callers can access the function.
- Place the code in the remote service.

## Next steps

Learn about [the various parameters](custom-functions-parameter-options.md) you can use in your custom functions. Or review the basics behind making [a web call through a custom function](custom-functions-web-reqs.md).

## See also

- [Volatile values in functions](custom-functions-volatile.md)
- [Create custom functions in Excel](custom-functions-overview.md)
- [Excel custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md)
