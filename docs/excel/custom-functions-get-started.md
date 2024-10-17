---
description: Get started with Excel custom functions for Office Add-ins.
title: Get started with custom functions in Excel
ms.date: 10/13/2024
ms.topic: overview
ms.custom: scenarios:getting-started
ms.localizationpriority: medium
---

# Get started with custom functions in Excel

This article includes tips, best practices, and Office Add-ins ecosystem information for new custom functions add-in developers.

The following diagram illustrates the interaction between a custom function and the two main components involved in custom function add-ins, Excel and external services.

:::image type="content" source="../images/custom-functions-add-in-components.png" alt-text="The custom functions add-in communicates with both Excel and an external service, but Excel and the external service don't communicate directly with each other.":::

**Excel** allows you to integrate your own custom functions into the application and run them like built-in functions.

The **custom functions add-in** defines the logic for your functions and how they interact with Excel and Office JavaScript APIs. To learn how to create a custom functions add-in, see the [custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md).

An **external service** is optional. It can give your add-in capabilities like importing data from outside the workbook. The custom functions add-in specifies how external data is incorporated into the workbook. To learn more, see [Receive and handle data with custom functions](custom-functions-web-reqs.md).

To develop a custom functions add-in with high performance, it’s important to ensure that these components operate in harmony. The following sections in this article describe how to optimize custom functions add-in development.

## Optimize custom functions recalculation efficiency

In general, custom functions recalculation follows the established pattern of [recalculation in Excel](/office/client-developer/excel/excel-recalculation). When recalculation is triggered, Excel enters a three-stage process: construct a dependency tree, construct a calculation chain, and then recalculate the cells. To optimize recalculation efficiency in your add-in, consider the level of nesting within your custom functions, the Excel calculation mode, and the limitations of volatile functions.

### Nesting in custom functions

A custom function can accept another custom function as an argument, ensuring that any dependent values are updated during recalculation. The recalculation of the outer custom function depends on the result of the nested function, leading to increased time consumption with each additional nested function. Minimize the number of nested levels in your custom functions to improve recalculation efficiency. The following code snippets demonstrate two approaches that produce similar outputs. **Option 1** is more likely to reduce the nested levels for end users when adding values in the workbook compared to **Option 2**.

#### Option 1

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

#### Option 2

```js
    /**
    * Returns the sum of two numbers.
    * @customfunction
    */
    function Add(arg1: number, arg2: number): number {
      return arg1 + arg2;
    }
```

### Excel calculation mode

Excel has three calculation modes: Automatic, Automatic Except Tables, and Manual. For a description of these calculation modes, see [Excel Recalculation](/office/client-developer/excel/excel-recalculation). The most frequently used calculation mode for custom functions add-ins is manual calculation mode. Set the appropriate calculation mode for your add-in with the [Excel.CalculationMode enum](/javascript/api/excel/excel.calculationmode) based on your scenario. Note that automatic calculation mode may trigger recalculation often and reduce the efficiency of your add-in.

### Volatile function limitations

Custom functions allow you to create your own volatile functions, similar to the `NOW` and `TODAY` functions in Excel. During recalculation, Excel evaluates cells that contain volatile functions and all of their dependent cells. As a result, using many volatile functions may make recalculation slow. Volatile functions should be used sparingly to optimize your custom functions add-in. For additional information, see [Volatile and Non-Volatile Functions](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).

## Design approaches to improve efficiency

Custom functions add-ins allow for flexible designs, which means that different add-in designs can produce the same output for your end users.

### Multiple results

You can return multiple results from your custom function with multiple functions or with one function.

To return multiple results with one function, use a dynamic array. This is usually the recommended approach because dynamic arrays only require updating a single cell to trigger recalculation for all results. To learn more about dynamic arrays in custom functions, see [Return multiple results from your custom function](custom-functions-dynamic-arrays.md).

Another way to return multiple results is to use multiple functions and return a single result for each function. A benefit of using multiple functions is that your end user can decide precisely which formula they want to update and then only trigger recalculation for that formula.

### Complex data structures

[Data types](custom-functions-data-types-concepts.md) are the best way to handle complex data structures in custom functions add-ins. Data types support [Excel errors](custom-functions-errors.md) and [formatted number values](custom-functions-data-types-concepts.md#output-a-formatted-number-value). Data types also allow for designing [entity value cards](excel-data-types-entity-card.md), extending Excel data beyond the 2-dimensional grid.

## Improve the user experience of remote data calls

Custom functions can fetch data from remote locations beyond the workbook, such as the web or a server. For more information about fetching remote data, see [Receive and handle data with custom functions](custom-functions-web-reqs.md). To maintain efficiency when making remote data calls, consider batching external calls, minimizing roundtrip duration for each call, and including messages in your add-in to communicate delays to your end user.

### Batch custom function remote calls

If your custom functions call a remote service, use a batching pattern to reduce the number of network calls to the remote service. To learn more, see [Batching custom function calls for a remote service](custom-functions-batching.md).

### Minimize roundtrip duration

Remote service connections can have a large impact on custom function performance. To reduce this impact, use these strategies:

- Server heavy processing should be handled in the remote server efficiently to shorten the end-to-end latency for a custom function. For example, have parallel computing designed on the server side. If your service is deployed on Azure, consider optimization using [high-performance computing on Azure](/azure/architecture/topics/high-performance-computing).
- Reduce the number of service calls by optimizing the add-in flow. For example, only send necessary calls to a remote service.

### Improve user-perceived performance through add-in UX

If a delay while calling a remote service is inevitable, consider providing end users with messages through add-in task pane to explain the delay. This gives the user information to help manage their expectations. The following image shows an example.

:::image type="content" source="../images/custom-functions-delay-example.png" alt-text="The delay message says 'It may take some time as we are getting the data ready for you'.":::

## See also

- [Receive and handle data with custom functions](custom-functions-web-reqs.md)
- [Batching custom function calls for a remote service](custom-functions-batching.md)
- [Custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md)
