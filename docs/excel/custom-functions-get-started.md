---
description: Tips and best practices for Excel custom functions in your Office Add-ins.
title: Best practices for custom functions in Excel
ms.date: 06/22/2025
ms.topic: overview
ms.custom: scenarios:getting-started
ms.localizationpriority: medium
---

# Best practices for custom functions in Excel

This article includes tips, best practices, and Office Add-ins ecosystem information for new developers of custom functions add-ins.

The following diagram illustrates the interaction between a custom function and the two main components involved in custom function add-ins: Excel and external services.

:::image type="content" source="../images/custom-functions-add-in-components.png" alt-text="The custom functions add-in communicates with both Excel and an external service, but Excel and the external service don't communicate directly with each other.":::

**Excel** allows you to integrate your own custom functions into the application and run them like built-in functions.

The **custom functions add-in** defines the logic for your functions and how they interact with Excel and Office JavaScript APIs. To learn how to create a custom functions add-in, see the [custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md).

An **external service** is optional. It can give your add-in capabilities like importing data from outside the workbook. The custom functions add-in specifies how external data is incorporated into the workbook. To learn more, see [Receive and handle data with custom functions](custom-functions-web-reqs.md).

## Optimize custom functions recalculation efficiency

In general, custom functions recalculation follows the established pattern of [recalculation in Excel](/office/client-developer/excel/excel-recalculation). When recalculation is triggered, Excel enters a three-stage process: construct a dependency tree, construct a calculation chain, and then recalculate the cells. To optimize recalculation efficiency in your add-in, consider the level of nesting within your custom functions, the Excel calculation mode, and the limitations of volatile functions.

### Nesting in custom functions

A custom function can accept another custom function as an argument, making the argument a nested custom function. The recalculation of the outer custom function depends on the result of the nested function, leading to increased time consumption with each additional nested function. Minimize the number of nested levels in your custom functions to improve recalculation efficiency. The following code snippets demonstrate two approaches for adding values in the workbook that produce similar outputs. **Option 1** uses an array to call values as a single parameter, while **Option 2** calls each value as a separate parameter, so **Option 1** is more efficient.

#### Option 1: Increase efficiency with limited nesting

> [!NOTE]
> This is the recommended approach. It uses an array to call values as a single parameter and avoid unnecessary nesting, so it's more efficient than **Option 2**.

```js
    /**
    * Returns the sum of input numbers.
    * @customfunction
    */
    function Add(args: number[]): number {
      let total = 0;
      args.forEach(value => {
        total += value;
      });
     
      return total;
    }
```

#### Option 2: More nesting is inefficient

> [!NOTE]
> This approach isn't recommended. **Option 1** and **Option 2** produce similar outputs, but **Option 2** uses more parameters and is less efficient.

```js
    /**
    * Returns the sum of two numbers.
    * @customfunction
    */
    function Add(arg1: number, arg2: number): number {
      return arg1 + arg2;
    }
```

### Excel calculation modes

Excel has three calculation modes: Automatic, Automatic Except Tables, and Manual. To determine which calculation mode best fits your custom function design, refer to the [Calculation Modes, Commands, Selective Recalculation, and Data Table](/office/client-developer/excel/excel-recalculation#calculation-modes-commands-selective-recalculation-and-data-tables#calculation-modes-commands-selective-recalculation-and-data-tables) section in the main [Excel Recalculation](/office/client-developer/excel/excel-recalculation) article.

Set the calculation mode for your add-in with the [Excel.CalculationMode enum](/javascript/api/excel/excel.calculationmode) based on your scenario. Note that `automatic` calculation mode may trigger recalculation often and reduce the efficiency of your add-in.

### Volatile function limitations

Custom functions allow you to create your own volatile functions, similar to the `NOW` and `TODAY` functions in Excel. During recalculation, Excel evaluates cells that contain volatile functions and all of their dependent cells. As a result, using many volatile functions may slow recalculation time so limit the number of volatile functions in your add-in to optimize efficiency. For additional information, see [Volatile and Non-Volatile Functions](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).

## Design approaches to improve efficiency

Custom functions add-ins allow for flexible designs, which means that different add-in designs can produce the same output for your end users.

### Multiple results

You can return multiple results from your custom function with multiple functions or with one function.

To return multiple results with one function, use a dynamic array. This is usually the recommended approach because dynamic arrays only require updating a single cell to trigger recalculation for all results.

:::image type="content" source="../images/custom-functions-dynamic-array.png" alt-text="The output of a dynamic array.":::

Keep in mind that using dynamic arrays becomes less efficient the larger your dataset is, because each recalculation processes more data. To learn more about dynamic arrays in custom functions, see [Return multiple results from your custom function](custom-functions-dynamic-arrays.md).

Another way to return multiple results is to use multiple functions and return a single result for each function. A benefit of using multiple functions is that your end user can decide precisely which formula they want to update and then only trigger recalculation for that formula. This is particularly helpful when relying on external services that may respond slowly.

:::image type="content" source="../images/custom-functions-not-dynamic-array.png" alt-text="The output of multiple functions instead of a dynamic array.":::

### Complex data structures

[Data types](custom-functions-data-types-concepts.md) are the best way to handle complex data structures in custom functions add-ins. Data types support [Excel errors](custom-functions-errors.md) and formatted numbers as [doubles](custom-functions-data-types-concepts.md#output-a-formatted-number). Data types also allow for designing [entity value cards](excel-data-types-entity-card.md), extending Excel data beyond the 2-dimensional grid.

## Improve the user experience of calls to external services

Custom functions can fetch data from remote locations beyond the workbook, such as the web or a server. For more information about fetching data from an external service, see [Receive and handle data with custom functions](custom-functions-web-reqs.md). To maintain efficiency when calling external services, consider batching external calls, minimizing roundtrip duration for each call, and including messages in your add-in to communicate delays to your end user.

### Batch custom function remote calls

If your custom functions call a remote service, use a batching pattern to reduce the number of network calls to the remote service. To learn more, see [Batching custom function calls for a remote service](custom-functions-batching.md).

### Minimize roundtrip duration

Remote service connections can have a large impact on custom function performance. To reduce this impact, use these strategies:

- Server-heavy processing should be handled efficiently in the remote server to shorten the end-to-end latency for a custom function. For example, have parallel computing designed on the server. If your service is deployed on Azure, consider optimization using [high-performance computing on Azure](/azure/architecture/topics/high-performance-computing).
- Reduce the number of service calls by optimizing the add-in flow. For example, only send necessary calls to a remote service.

### Improve user-perceived performance through add-in user experience (UX)

While a custom function is calling an external service, the cell with the custom function displays the **#BUSY!** error. If a delay while calling an external service is inevitable, consider providing messages through the add-in task pane to explain the delay to your end users. This information helps manage their expectations. The following image shows an example.

:::image type="content" source="../images/custom-functions-delay-example.png" alt-text="The delay message says 'There may be a delay. We're getting the data ready for you'.":::

For more information about how to share data between a custom function and a task pane, see [Share data and events between Excel custom functions and the task pane](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md).

To display a message in the add-in task pane that notifies users of a delay, make the following changes after ensuring that your add-in uses a shared runtime.

1. In **taskpane.js** add a function that calls the notification.

    ```js
    export function showNotification(message){
      const label = document.getElementById("item-subject");
      label.innerHTML = message;
    }
    ```

1. In **function.js**, import the `showNotification` function.

    ```js
    export function showNotification(message){
      const label = document.getElementById("item-subject");
      label.innerHTML = message;
    }
    ```

1. In **function.js**, call `showNotification` when running the calculation that may include a delay.

    ```js
    export async function longCalculation(param) {
      await Office.addin.showAsTaskpane();
      showNotification("It may take some time as we prepare the data.");
      // Perform long operation
      // ...
      // ...
      return answer;
    }
    ```

## See also

- [Receive and handle data with custom functions](custom-functions-web-reqs.md)
- [Batch custom function calls for a remote service](custom-functions-batching.md)
- [Create custom functions in Excel tutorial](../tutorials/excel-tutorial-create-custom-functions.md)
